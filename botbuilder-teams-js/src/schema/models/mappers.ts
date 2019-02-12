/*
 * Code generated by Microsoft (R) AutoRest Code Generator.
 * Changes may cause incorrect behavior and will be lost if the code is
 * regenerated.
 */

import * as msRest from 'botframework-connector/node_modules/ms-rest-js';


export const ChannelInfo: msRest.CompositeMapper = {
  serializedName: 'ChannelInfo',
  type: {
    name: 'Composite',
    className: 'ChannelInfo',
    modelProperties: {
      id: {
        serializedName: 'id',
        type: {
          name: 'String'
        }
      },
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const ConversationList: msRest.CompositeMapper = {
  serializedName: 'ConversationList',
  type: {
    name: 'Composite',
    className: 'ConversationList',
    modelProperties: {
      conversations: {
        serializedName: 'conversations',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'ChannelInfo'
            }
          }
        }
      }
    }
  }
};

export const TeamDetails: msRest.CompositeMapper = {
  serializedName: 'TeamDetails',
  type: {
    name: 'Composite',
    className: 'TeamDetails',
    modelProperties: {
      id: {
        serializedName: 'id',
        type: {
          name: 'String'
        }
      },
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      },
      aadGroupId: {
        serializedName: 'aadGroupId',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const TeamInfo: msRest.CompositeMapper = {
  serializedName: 'TeamInfo',
  type: {
    name: 'Composite',
    className: 'TeamInfo',
    modelProperties: {
      id: {
        serializedName: 'id',
        type: {
          name: 'String'
        }
      },
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const NotificationInfo: msRest.CompositeMapper = {
  serializedName: 'NotificationInfo',
  type: {
    name: 'Composite',
    className: 'NotificationInfo',
    modelProperties: {
      alert: {
        serializedName: 'alert',
        type: {
          name: 'Boolean'
        }
      }
    }
  }
};

export const TenantInfo: msRest.CompositeMapper = {
  serializedName: 'TenantInfo',
  type: {
    name: 'Composite',
    className: 'TenantInfo',
    modelProperties: {
      id: {
        serializedName: 'id',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const TeamsChannelData: msRest.CompositeMapper = {
  serializedName: 'TeamsChannelData',
  type: {
    name: 'Composite',
    className: 'TeamsChannelData',
    modelProperties: {
      channel: {
        serializedName: 'channel',
        type: {
          name: 'Composite',
          className: 'ChannelInfo'
        }
      },
      eventType: {
        serializedName: 'eventType',
        type: {
          name: 'String'
        }
      },
      team: {
        serializedName: 'team',
        type: {
          name: 'Composite',
          className: 'TeamInfo'
        }
      },
      notification: {
        serializedName: 'notification',
        type: {
          name: 'Composite',
          className: 'NotificationInfo'
        }
      },
      tenant: {
        serializedName: 'tenant',
        type: {
          name: 'Composite',
          className: 'TenantInfo'
        }
      }
    }
  }
};

export const ChannelAccount: msRest.CompositeMapper = {
  serializedName: 'ChannelAccount',
  type: {
    name: 'Composite',
    className: 'ChannelAccount',
    modelProperties: {
      id: {
        serializedName: 'id',
        type: {
          name: 'String'
        }
      },
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const TeamsChannelAccount: msRest.CompositeMapper = {
  serializedName: 'TeamsChannelAccount',
  type: {
    name: 'Composite',
    className: 'TeamsChannelAccount',
    modelProperties: {
      ...ChannelAccount.type.modelProperties,
      aadObjectId: {
        serializedName: 'aadObjectId',
        type: {
          name: 'String'
        }
      },
      givenName: {
        serializedName: 'givenName',
        type: {
          name: 'String'
        }
      },
      surname: {
        serializedName: 'surname',
        type: {
          name: 'String'
        }
      },
      email: {
        serializedName: 'email',
        type: {
          name: 'String'
        }
      },
      userPrincipalName: {
        serializedName: 'userPrincipalName',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const CardAction: msRest.CompositeMapper = {
  serializedName: 'CardAction',
  type: {
    name: 'Composite',
    className: 'CardAction',
    modelProperties: {
      type: {
        serializedName: 'type',
        type: {
          name: 'String'
        }
      },
      title: {
        serializedName: 'title',
        type: {
          name: 'String'
        }
      },
      image: {
        serializedName: 'image',
        type: {
          name: 'String'
        }
      },
      value: {
        serializedName: 'value',
        type: {
          name: 'Object'
        }
      }
    }
  }
};

export const CardImage: msRest.CompositeMapper = {
  serializedName: 'CardImage',
  type: {
    name: 'Composite',
    className: 'CardImage',
    modelProperties: {
      url: {
        serializedName: 'url',
        type: {
          name: 'String'
        }
      },
      alt: {
        serializedName: 'alt',
        type: {
          name: 'String'
        }
      },
      tap: {
        serializedName: 'tap',
        type: {
          name: 'Composite',
          className: 'CardAction'
        }
      }
    }
  }
};

export const O365ConnectorCardFact: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardFact',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardFact',
    modelProperties: {
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      },
      value: {
        serializedName: 'value',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardImage: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardImage',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardImage',
    modelProperties: {
      image: {
        serializedName: 'image',
        type: {
          name: 'String'
        }
      },
      title: {
        serializedName: 'title',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardActionBase: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardActionBase',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardActionBase',
    modelProperties: {
      type: {
        serializedName: '@type',
        type: {
          name: 'String'
        }
      },
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      },
      id: {
        serializedName: '@id',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardSection: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardSection',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardSection',
    modelProperties: {
      title: {
        serializedName: 'title',
        type: {
          name: 'String'
        }
      },
      text: {
        serializedName: 'text',
        type: {
          name: 'String'
        }
      },
      activityTitle: {
        serializedName: 'activityTitle',
        type: {
          name: 'String'
        }
      },
      activitySubtitle: {
        serializedName: 'activitySubtitle',
        type: {
          name: 'String'
        }
      },
      activityText: {
        serializedName: 'activityText',
        type: {
          name: 'String'
        }
      },
      activityImage: {
        serializedName: 'activityImage',
        type: {
          name: 'String'
        }
      },
      activityImageType: {
        serializedName: 'activityImageType',
        type: {
          name: 'String'
        }
      },
      markdown: {
        serializedName: 'markdown',
        type: {
          name: 'Boolean'
        }
      },
      facts: {
        serializedName: 'facts',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardFact'
            }
          }
        }
      },
      images: {
        serializedName: 'images',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardImage'
            }
          }
        }
      },
      potentialAction: {
        serializedName: 'potentialAction',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardActionBase'
            }
          }
        }
      }
    }
  }
};

export const O365ConnectorCard: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCard',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCard',
    modelProperties: {
      title: {
        serializedName: 'title',
        type: {
          name: 'String'
        }
      },
      text: {
        serializedName: 'text',
        type: {
          name: 'String'
        }
      },
      summary: {
        serializedName: 'summary',
        type: {
          name: 'String'
        }
      },
      themeColor: {
        serializedName: 'themeColor',
        type: {
          name: 'String'
        }
      },
      sections: {
        serializedName: 'sections',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardSection'
            }
          }
        }
      },
      potentialAction: {
        serializedName: 'potentialAction',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardActionBase'
            }
          }
        }
      }
    }
  }
};

export const O365ConnectorCardViewAction: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardViewAction',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardViewAction',
    modelProperties: {
      ...O365ConnectorCardActionBase.type.modelProperties,
      target: {
        serializedName: 'target',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'String'
            }
          }
        }
      }
    }
  }
};

export const O365ConnectorCardOpenUriTarget: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardOpenUriTarget',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardOpenUriTarget',
    modelProperties: {
      os: {
        serializedName: 'os',
        type: {
          name: 'String'
        }
      },
      uri: {
        serializedName: 'uri',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardOpenUri: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardOpenUri',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardOpenUri',
    modelProperties: {
      ...O365ConnectorCardActionBase.type.modelProperties,
      targets: {
        serializedName: 'targets',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardOpenUriTarget'
            }
          }
        }
      }
    }
  }
};

export const O365ConnectorCardHttpPOST: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardHttpPOST',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardHttpPOST',
    modelProperties: {
      ...O365ConnectorCardActionBase.type.modelProperties,
      body: {
        serializedName: 'body',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardInputBase: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardInputBase',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardInputBase',
    modelProperties: {
      type: {
        serializedName: '@type',
        type: {
          name: 'String'
        }
      },
      id: {
        serializedName: 'id',
        type: {
          name: 'String'
        }
      },
      isRequired: {
        serializedName: 'isRequired',
        type: {
          name: 'Boolean'
        }
      },
      title: {
        serializedName: 'title',
        type: {
          name: 'String'
        }
      },
      value: {
        serializedName: 'value',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardActionCard: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardActionCard',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardActionCard',
    modelProperties: {
      ...O365ConnectorCardActionBase.type.modelProperties,
      inputs: {
        serializedName: 'inputs',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardInputBase'
            }
          }
        }
      },
      actions: {
        serializedName: 'actions',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardActionBase'
            }
          }
        }
      }
    }
  }
};

export const O365ConnectorCardTextInput: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardTextInput',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardTextInput',
    modelProperties: {
      ...O365ConnectorCardInputBase.type.modelProperties,
      isMultiline: {
        serializedName: 'isMultiline',
        type: {
          name: 'Boolean'
        }
      },
      maxLength: {
        serializedName: 'maxLength',
        type: {
          name: 'Number'
        }
      }
    }
  }
};

export const O365ConnectorCardDateInput: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardDateInput',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardDateInput',
    modelProperties: {
      ...O365ConnectorCardInputBase.type.modelProperties,
      includeTime: {
        serializedName: 'includeTime',
        type: {
          name: 'Boolean'
        }
      }
    }
  }
};

export const O365ConnectorCardMultichoiceInputChoice: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardMultichoiceInputChoice',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardMultichoiceInputChoice',
    modelProperties: {
      display: {
        serializedName: 'display',
        type: {
          name: 'String'
        }
      },
      value: {
        serializedName: 'value',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const O365ConnectorCardMultichoiceInput: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardMultichoiceInput',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardMultichoiceInput',
    modelProperties: {
      ...O365ConnectorCardInputBase.type.modelProperties,
      choices: {
        serializedName: 'choices',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'O365ConnectorCardMultichoiceInputChoice'
            }
          }
        }
      },
      style: {
        serializedName: 'style',
        type: {
          name: 'String'
        }
      },
      isMultiSelect: {
        serializedName: 'isMultiSelect',
        type: {
          name: 'Boolean'
        }
      }
    }
  }
};

export const O365ConnectorCardActionQuery: msRest.CompositeMapper = {
  serializedName: 'O365ConnectorCardActionQuery',
  type: {
    name: 'Composite',
    className: 'O365ConnectorCardActionQuery',
    modelProperties: {
      body: {
        serializedName: 'body',
        type: {
          name: 'String'
        }
      },
      actionId: {
        serializedName: 'actionId',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const SigninStateVerificationQuery: msRest.CompositeMapper = {
  serializedName: 'SigninStateVerificationQuery',
  type: {
    name: 'Composite',
    className: 'SigninStateVerificationQuery',
    modelProperties: {
      state: {
        serializedName: 'state',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const MessagingExtensionQueryOptions: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionQueryOptions',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionQueryOptions',
    modelProperties: {
      skip: {
        serializedName: 'skip',
        type: {
          name: 'Number'
        }
      },
      count: {
        serializedName: 'count',
        type: {
          name: 'Number'
        }
      }
    }
  }
};

export const MessagingExtensionParameter: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionParameter',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionParameter',
    modelProperties: {
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      },
      value: {
        serializedName: 'value',
        type: {
          name: 'Object'
        }
      }
    }
  }
};

export const MessagingExtensionQuery: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionQuery',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionQuery',
    modelProperties: {
      commandId: {
        serializedName: 'commandId',
        type: {
          name: 'String'
        }
      },
      parameters: {
        serializedName: 'parameters',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'MessagingExtensionParameter'
            }
          }
        }
      },
      queryOptions: {
        serializedName: 'queryOptions',
        type: {
          name: 'Composite',
          className: 'MessagingExtensionQueryOptions'
        }
      },
      state: {
        serializedName: 'state',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const Attachment: msRest.CompositeMapper = {
  serializedName: 'Attachment',
  type: {
    name: 'Composite',
    className: 'Attachment',
    modelProperties: {
      contentType: {
        serializedName: 'contentType',
        type: {
          name: 'String'
        }
      },
      contentUrl: {
        serializedName: 'contentUrl',
        type: {
          name: 'String'
        }
      },
      content: {
        serializedName: 'content',
        type: {
          name: 'Object'
        }
      },
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      },
      thumbnailUrl: {
        serializedName: 'thumbnailUrl',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const MessagingExtensionAttachment: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionAttachment',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionAttachment',
    modelProperties: {
      ...Attachment.type.modelProperties,
      preview: {
        serializedName: 'preview',
        type: {
          name: 'Composite',
          className: 'Attachment'
        }
      }
    }
  }
};

export const MessagingExtensionSuggestedAction: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionSuggestedAction',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionSuggestedAction',
    modelProperties: {
      actions: {
        serializedName: 'actions',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'CardAction'
            }
          }
        }
      }
    }
  }
};

export const MessagingExtensionResult: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionResult',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionResult',
    modelProperties: {
      attachmentLayout: {
        serializedName: 'attachmentLayout',
        type: {
          name: 'String'
        }
      },
      type: {
        serializedName: 'type',
        type: {
          name: 'String'
        }
      },
      attachments: {
        serializedName: 'attachments',
        type: {
          name: 'Sequence',
          element: {
            type: {
              name: 'Composite',
              className: 'MessagingExtensionAttachment'
            }
          }
        }
      },
      suggestedActions: {
        serializedName: 'suggestedActions',
        type: {
          name: 'Composite',
          className: 'MessagingExtensionSuggestedAction'
        }
      },
      text: {
        serializedName: 'text',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const MessagingExtensionResponse: msRest.CompositeMapper = {
  serializedName: 'MessagingExtensionResponse',
  type: {
    name: 'Composite',
    className: 'MessagingExtensionResponse',
    modelProperties: {
      composeExtension: {
        serializedName: 'composeExtension',
        type: {
          name: 'Composite',
          className: 'MessagingExtensionResult'
        }
      }
    }
  }
};

export const FileConsentCard: msRest.CompositeMapper = {
  serializedName: 'FileConsentCard',
  type: {
    name: 'Composite',
    className: 'FileConsentCard',
    modelProperties: {
      description: {
        serializedName: 'description',
        type: {
          name: 'String'
        }
      },
      sizeInBytes: {
        serializedName: 'sizeInBytes',
        type: {
          name: 'Number'
        }
      },
      acceptContext: {
        serializedName: 'acceptContext',
        type: {
          name: 'Object'
        }
      },
      declineContext: {
        serializedName: 'declineContext',
        type: {
          name: 'Object'
        }
      }
    }
  }
};

export const FileDownloadInfo: msRest.CompositeMapper = {
  serializedName: 'FileDownloadInfo',
  type: {
    name: 'Composite',
    className: 'FileDownloadInfo',
    modelProperties: {
      downloadUrl: {
        serializedName: 'downloadUrl',
        type: {
          name: 'String'
        }
      },
      uniqueId: {
        serializedName: 'uniqueId',
        type: {
          name: 'String'
        }
      },
      fileType: {
        serializedName: 'fileType',
        type: {
          name: 'String'
        }
      },
      etag: {
        serializedName: 'etag',
        type: {
          name: 'Object'
        }
      }
    }
  }
};

export const FileInfoCard: msRest.CompositeMapper = {
  serializedName: 'FileInfoCard',
  type: {
    name: 'Composite',
    className: 'FileInfoCard',
    modelProperties: {
      uniqueId: {
        serializedName: 'uniqueId',
        type: {
          name: 'String'
        }
      },
      fileType: {
        serializedName: 'fileType',
        type: {
          name: 'String'
        }
      },
      etag: {
        serializedName: 'etag',
        type: {
          name: 'Object'
        }
      }
    }
  }
};

export const FileUploadInfo: msRest.CompositeMapper = {
  serializedName: 'FileUploadInfo',
  type: {
    name: 'Composite',
    className: 'FileUploadInfo',
    modelProperties: {
      name: {
        serializedName: 'name',
        type: {
          name: 'String'
        }
      },
      uploadUrl: {
        serializedName: 'uploadUrl',
        type: {
          name: 'String'
        }
      },
      contentUrl: {
        serializedName: 'contentUrl',
        type: {
          name: 'String'
        }
      },
      uniqueId: {
        serializedName: 'uniqueId',
        type: {
          name: 'String'
        }
      },
      fileType: {
        serializedName: 'fileType',
        type: {
          name: 'String'
        }
      }
    }
  }
};

export const FileConsentCardResponse: msRest.CompositeMapper = {
  serializedName: 'FileConsentCardResponse',
  type: {
    name: 'Composite',
    className: 'FileConsentCardResponse',
    modelProperties: {
      action: {
        serializedName: 'action',
        type: {
          name: 'String'
        }
      },
      context: {
        serializedName: 'context',
        type: {
          name: 'Object'
        }
      },
      uploadInfo: {
        serializedName: 'uploadInfo',
        type: {
          name: 'Composite',
          className: 'FileUploadInfo'
        }
      }
    }
  }
};
