// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import { TurnContext, InvokeResponse, Activity } from 'botbuilder';
import { 
  MessagingExtensionQuery, 
  MessagingExtensionResponse,
  O365ConnectorCardActionQuery,
  SigninStateVerificationQuery,
  FileConsentCardResponse,
  TaskModuleRequest,
  TaskModuleResponse,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  AppBasedLinkQuery
} from '../schema';

/**
 * Typed invoke request activity, inherited from `Activity`
 * @typeparam T the type to define invoke `value` free payload
 */
export interface InvokeRequestActivity<T> extends Activity {
  value: T;
}

/**
 * Typed invoke response, inherited from `InvokeResponse`
 * @typeparam T the type to define `body` free payload in response
 */
export interface InvokeResponseTyped<T> extends InvokeResponse {
  body: T;
}

export type InvokeType = keyof typeof InvokeActivity.definitions;
export type InvokeValueTypeOf<N extends InvokeType> = (typeof InvokeActivity.definitions)[N]['value'];
export type InvokeResponseUnsafeTypeOf<N extends InvokeType> = (typeof InvokeActivity.definitions)[N]['response'];
export type InvokeResponseTypeOf<N extends InvokeType> = InvokeResponseUnsafeTypeOf<N> extends InvokeResponse ? InvokeResponseUnsafeTypeOf<N>: InvokeResponse;

/**
 * Define type-binding event handlers for Teams invoke activities.
 * 
 * @remarks 
 * This type definition defines a list of interfaces for type-binding invoke handlers, 
 * with typed interface name, invoke value (request), and invoke response whose definitions are populated from `InvokeActivity.definitions`
 * This is also the base class of `ITeamsInvokeActivityHandler`
 */
export type InvokeTypedHandler = {
  [name in InvokeType]?: (turnContext: TurnContext, invokeValue: InvokeValueTypeOf<name>) => Promise<InvokeResponseTypeOf<name>>;
};

/**
 * Event handlers for Teams invoke activities.
 * 
 * @remarks
 * This class extends the type of `InvokeTypedHandler` that defines type-binding invoke request and response
 * where the definitions are given by `InvokeActivity.definitions`. In consequence, `ITeamsInvokeActivityHandler`
 * will actually define a list of interface for type-binding invoke handlers, plus a generic `onInvoke` handler
 */
export interface ITeamsInvokeActivityHandler extends InvokeTypedHandler {
  /**
   * Handles generic invoke request. This handler will be triggered only when none of type-binding invoke handlers is applied. 
   * @param turnContext Current turn context.
   */
  onInvoke? (turnContext: TurnContext): Promise<InvokeResponse>;
}

/**
 * Helper class for invoke activity
 */
export class InvokeActivity {
  /**
   * List of type-binding invoke handler definitions
   * @remarks
   * The key used in definitions will turn out to be the name of invoke handler. Similarly, for the values:
   * `name` defines the name of invoke to trigger this handler;
   * `value` defines the invoke request value type;
   * `response` by default it's `InvokeResponse` (no type-binding response). If response body is typed, it must inherit from `InvokeResponse`
   *  an easier way is to use `InvokeResponseTyped<T>` to wrap `T` into typed invoke response. Note that if arbitrary type assigned 
   *  without inheriting `InvokeResponse`, it'll be overwritten to be `InvokeResponse` by our auto type deductions.
   */
  public static readonly definitions = {
    onO365CardAction: {
      name: 'actionableMessage/executeAction',
      value: <O365ConnectorCardActionQuery> {},
      response: <InvokeResponse>{}
    },

    onSigninStateVerification: {
      name: 'signin/verifyState',
      value: <SigninStateVerificationQuery> {},
      response: <InvokeResponse>{}
    },

    onFileConsent: {
      name: 'fileConsent/invoke',
      value: <FileConsentCardResponse> {},
      response: <InvokeResponse> {}
    },

    onMessagingExtensionQuery: {
      name: 'composeExtension/query',
      value: <MessagingExtensionQuery> {},
      response: <InvokeResponseTyped<MessagingExtensionResponse>> {}
    },

    onAppBasedLinkQuery: {
      name: 'composeExtension/queryLink',
      value: <AppBasedLinkQuery> {},
      response: <InvokeResponseTyped<MessagingExtensionResponse>> {}
    },

    onMessagingExtensionFetchTask: {
      name: 'composeExtension/fetchTask',
      value: <MessagingExtensionAction> {},
      response: <InvokeResponseTyped<MessagingExtensionActionResponse>> {}
    },

    onMessagingExtensionSubmitAction: {
      name: 'composeExtension/submitAction',
      value: <MessagingExtensionAction> {},
      response: <InvokeResponseTyped<MessagingExtensionActionResponse>> {}
    },

    onTaskModuleFetch: {
      name: 'task/fetch',
      value: <TaskModuleRequest> {},
      response: <InvokeResponseTyped<TaskModuleResponse>> {}
    },

    onTaskModuleSubmit: {
      name: 'task/submit',
      value: <TaskModuleRequest> {},
      response: <InvokeResponseTyped<TaskModuleResponse>> {}
    }
  };

  /**
   * Type guard for activity auto casting into one with typed "value" for invoke payload
   * @param {Activity} activity Activity
   * @param N InvokeType, i.e., the key names of InvokeActivity.definitions
   */
  public static is<N extends InvokeType>(activity: Activity, N: N):
    activity is InvokeRequestActivity<InvokeValueTypeOf<N>> {
    return activity.name === InvokeActivity.definitions[N].name;
  }

  /**
   * Dispatch invoke event to corresponding invoke handlers
   * @param handler the handler set of invoke events
   * @param turnContext Current turn context.
   */
  public static async dispatchHandler(handler: ITeamsInvokeActivityHandler, turnContext: TurnContext) {
    if (handler) {
      const activity = turnContext.activity;

      if (handler.onO365CardAction && InvokeActivity.is(activity, 'onO365CardAction')) {
        return await handler.onO365CardAction(turnContext, activity.value);
      }

      if (handler.onSigninStateVerification && InvokeActivity.is(activity, 'onSigninStateVerification')) {
        return await handler.onSigninStateVerification(turnContext, activity.value);
      }

      if (handler.onFileConsent && InvokeActivity.is(activity, 'onFileConsent')) {
        return await handler.onFileConsent(turnContext, activity.value);
      }

      if (handler.onMessagingExtensionQuery && InvokeActivity.is(activity, 'onMessagingExtensionQuery')) {
        return await handler.onMessagingExtensionQuery(turnContext, activity.value);
      }

      if (handler.onAppBasedLinkQuery && InvokeActivity.is(activity, 'onAppBasedLinkQuery')) {
        return await handler.onAppBasedLinkQuery(turnContext, activity.value);
      }

      if (handler.onMessagingExtensionFetchTask && InvokeActivity.is(activity, 'onMessagingExtensionFetchTask')) {
        return await handler.onMessagingExtensionFetchTask(turnContext, activity.value);
      }

      if (handler.onMessagingExtensionSubmitAction && InvokeActivity.is(activity, 'onMessagingExtensionSubmitAction')) {
        return await handler.onMessagingExtensionSubmitAction(turnContext, activity.value);
      }

      if (handler.onTaskModuleFetch && InvokeActivity.is(activity, 'onTaskModuleFetch')) {
        return await handler.onTaskModuleFetch(turnContext, activity.value);
      }

      if (handler.onTaskModuleSubmit && InvokeActivity.is(activity, 'onTaskModuleSubmit')) {
        return await handler.onTaskModuleSubmit(turnContext, activity.value);
      }

      if (handler.onInvoke) {
        return await handler.onInvoke(turnContext);
      }
    }
  }
}
