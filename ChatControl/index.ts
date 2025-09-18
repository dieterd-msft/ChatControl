import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom"; 
import Chat, { IChatProps, ChatRef } from "./Chat/Chat";

import {
  ConnectionSettings,
  CopilotStudioClient,
} from "@microsoft/agents-copilotstudio-client";


export class ChatControl
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  // Element id of the ng-app div. Type: string
  private _appDivId: string;
  private _appDiv: HTMLDivElement;

  // ng-app app id. Type: string
  private _appId: string;

  // ng-controller. Type: string
  private _controllerId: string;

  // PCF framework delegate which will be assigned to this object which would be called whenever any update happens. Type: function
  private _notifyOutputChanged: () => void;

  private _context: ComponentFramework.Context<IInputs>;

  private _settings: IChatProps;

  private _container: HTMLDivElement;
  private _inputElement: React.ReactElement;
  private _chatRef = React.createRef<ChatRef>();

  private _prevMessage: string | undefined;
  private _agentMessage: string | undefined;
  private _chatKey = 1;

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._settings = {
      appClientId: context.parameters.appClientId.raw || "",
      tenantId: context.parameters.tenantId.raw || "",
      currentUserLogin: context.parameters.username.raw || "",
      baseUrl: window.location.origin + "/",
      environmentId: context.parameters.environmentId.raw || "",
      agentIdentifier: context.parameters.agentIdentifier.raw || "",
      directConnectUrl:
        "https://" +
          this.convertUUID(context.parameters.environmentId.raw || "") +
          ".environment.api.powerplatform.com/copilotstudio/dataverse-backed/authenticated/bots/" +
          context.parameters.agentIdentifier.raw ||
        "" + "/conversations?api-version=2022-03-01-preview",
      styleOptions: context.parameters.styleOptions.raw || "",
    };
    this._container = container;
    context.mode.trackContainerResize(true);

    this._appDiv = document.createElement("div");
    this._appDiv.setAttribute("id", this._appDivId);
    this._appDiv.setAttribute("ng-controller", this._appId);
    this._appDiv.setAttribute("ng-app", this._controllerId);
    
    const allocatedWidth = context.mode.allocatedWidth;
        // Use allocatedWidth to set the component's width dynamically
    this._appDiv.style.width = allocatedWidth + "px";
    const allocatedHeight = context.mode.allocatedHeight;
    // Use allocatedHeight to set the component's height dynamically
    this._appDiv.style.height = allocatedHeight + "px";

    this._notifyOutputChanged = notifyOutputChanged;

    container.appendChild(this._appDiv);
    this._inputElement = React.createElement(Chat, {
      key: this._chatKey,
      appClientId: this._settings.appClientId,
      tenantId: this._settings.tenantId,
      environmentId: this._settings.environmentId,
      agentIdentifier: this._settings.agentIdentifier,
      directConnectUrl: this._settings.directConnectUrl,
      showTyping: true,
      currentUserLogin: this._settings.currentUserLogin,
      baseUrl: this._settings.baseUrl,
      styleOptions: this._settings.styleOptions,
      width: allocatedWidth+'px',
      height: allocatedHeight+'px',
      onAgentMessageUpdate: (message: string) => {
        this._agentMessage = message;
        this._notifyOutputChanged();
      },
      onNewConversation: this._startNewConversation,
      ref: this._chatRef,
    });
    ReactDOM.render(this._inputElement, this._appDiv, () => {
      console.log("Chat component rendered successfully");
    });
  }
  private _startNewConversation = (): void => {
    this._chatKey += 1;
    this._recreateChatComponent();
  };

  private _recreateChatComponent(): void {
    const allocatedWidth = this._context.mode.allocatedWidth;
    const allocatedHeight = this._context.mode.allocatedHeight;
    
    this._inputElement = React.createElement(Chat, {
      key: this._chatKey,
      appClientId: this._settings.appClientId,
      tenantId: this._settings.tenantId,
      environmentId: this._settings.environmentId,
      agentIdentifier: this._settings.agentIdentifier,
      directConnectUrl: this._settings.directConnectUrl,
      showTyping: true,
      currentUserLogin: this._settings.currentUserLogin,
      baseUrl: this._settings.baseUrl,
      styleOptions: this._settings.styleOptions,
      width: allocatedWidth+'px',
      height: allocatedHeight+'px',
      onAgentMessageUpdate: (message: string) => {
        this._agentMessage = message;
        this._notifyOutputChanged();
      },
      onNewConversation: this._startNewConversation,
      ref: this._chatRef,
    });

    ReactDOM.render(this._inputElement, this._appDiv);
  };


  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Only call sendMessage if message changed
    const newMessage = context.parameters.message?.raw ?? undefined;
    if ((newMessage != "val" || newMessage != null) && newMessage !== this._prevMessage) {
      this._prevMessage = newMessage;
      if (this._chatRef.current) {
        this._chatRef.current.sendMessage(context.parameters.message?.raw || "");
      }
    } else {
      const allocatedWidth = context.mode.allocatedWidth;
      const allocatedHeight = context.mode.allocatedHeight;
      
      this._inputElement = React.createElement(Chat, {
        appClientId: this._settings.appClientId,
        tenantId: this._settings.tenantId,
        environmentId: this._settings.environmentId,
        agentIdentifier: this._settings.agentIdentifier,
        directConnectUrl: this._settings.directConnectUrl,
        showTyping: true,
        currentUserLogin: this._settings.currentUserLogin,
        baseUrl: this._settings.baseUrl,
        styleOptions: this._settings.styleOptions,
        width: allocatedWidth+'px',
        height: allocatedHeight+'px',
        onAgentMessageUpdate: (message: string) => {
          this._agentMessage = message;
          this._notifyOutputChanged();
        },
        onNewConversation: this._startNewConversation,
        ref: this._chatRef,
      });
      ReactDOM.render(this._inputElement, this._appDiv, () => {
        console.log("Chat component rendered successfully");
      });
    }
    return;
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    return {
      response: this._agentMessage,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    if (this._container) {
      ReactDOM.unmountComponentAtNode(this._container);
    }
  }

  public convertUUID(uuid: string): string {
    // Remove hyphens
    const cleaned = uuid.replace(/-/g, "");

    // Split into two parts: first 30 characters and last 2 characters
    const mainPart = cleaned.slice(0, 30);
    const decimalPart = cleaned.slice(30);

    // Combine with a dot
    return `${mainPart}.${decimalPart}`;
  }
}
