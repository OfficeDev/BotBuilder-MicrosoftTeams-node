import { TurnContext, InvokeResponse, Activity } from 'botbuilder';
import { MessagingExtensionQuery, MessagingExtensionResponse, O365ConnectorCardActionQuery, SigninStateVerificationQuery, FileConsentCardResponse } from '../schema';

export interface IInvokeActivity<T> extends Activity {
  value: T;
}

export interface IInvokeResponseTyped<T> extends InvokeResponse {
  body: T;
}

export type InvokeType = keyof typeof InvokeActivity.definitions;
export type InvokeValueTypeOf<N extends InvokeType> = (typeof InvokeActivity.definitions)[N]['value'];
export type InvokeResponseUnsafeTypeOf<N extends InvokeType> = (typeof InvokeActivity.definitions)[N]['response'];
export type InvokeResponseTypeOf<N extends InvokeType> = InvokeResponseUnsafeTypeOf<N> extends InvokeResponse ? InvokeResponseUnsafeTypeOf<N>: InvokeResponse;
export type InvokeTypedHandler = {
  [name in InvokeType]?: (turnContext: TurnContext, invokeValue: InvokeValueTypeOf<name>) => Promise<InvokeResponseTypeOf<name>>;
};

export interface ITeamsInvokeActivityHandler extends InvokeTypedHandler {
  onInvoke? (turnContext: TurnContext): Promise<InvokeResponse>;
}

export class InvokeActivity {
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
      response: <IInvokeResponseTyped<MessagingExtensionResponse>> {}
    },
  };

  /**
   * Type guard for activity auto casting into one with typed "value" for invoke payload
   * @param activity Activity
   * @param N InvokeType, i.e., the key names of InvokeActivity.definitions
   */
  public static is<N extends InvokeType>(activity: Activity, N: N):
    activity is IInvokeActivity<InvokeValueTypeOf<N>> {
    return activity.name === InvokeActivity.definitions[N].name;
  }

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

      if (handler.onInvoke) {
        return await handler.onInvoke(turnContext);
      }
    }
  }
}
