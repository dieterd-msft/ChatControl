import * as React from "react";
import { useEffect, useState } from "react";
import { Components, createStore, ReactWebChat } from "botframework-webchat";
import { FluentThemeProvider } from "botframework-webchat-fluent-theme";

import {
  ConnectionSettings,
  CopilotStudioClient,
  CopilotStudioWebChat,
  CopilotStudioWebChatConnection,
} from "@microsoft/agents-copilotstudio-client";
import { acquireToken } from "./acquireToken";
import { on } from "events";
import { send } from "process";

const { BasicWebChat, Composer } = Components;

export interface IChatProps {
  chatKey?: number;
  agentTitle?: string;
  appClientId: string;
  tenantId: string;
  environmentId?: string;
  agentIdentifier?: string;
  directConnectUrl?: string;
  showTyping?: boolean;
  disableFileUploadButton?: boolean;
  disableHeader?: boolean;
  currentUserLogin?: string;
  baseUrl?: string;
  styleOptions?: string;
  width?: string | number;
  height?: string | number;
  onAgentMessageUpdate?: (message: string) => void;
  onConversationIdUpdate?: (conversationId: string) => void;
  onNewConversation?: () => void;
}

export interface ChatRef {
  sendMessage: (text: string) => void;
  sendEvent: (name: string, value?: unknown) => void;
  updateAgentMessage: (message: string) => void;
}

const Chat = React.forwardRef<ChatRef, IChatProps>(
  (
    {
      agentTitle,
      appClientId,
      tenantId,
      environmentId,
      agentIdentifier,
      directConnectUrl,
      showTyping = true,
      disableFileUploadButton,
      disableHeader,
      currentUserLogin,
      baseUrl,
      width,
      height,
      styleOptions,
      onAgentMessageUpdate,
      onConversationIdUpdate,
      onNewConversation,
    },
    ref
  ) => {
    const [connection, setConnection] =
      useState<CopilotStudioWebChatConnection | null>(null);
    const [error, setError] = useState<string>();
    const [store] = useState(() => createStore());
    const [lastConversationId, setLastConversationId] = useState<string | null>(
      null
    );

    // Set up store listener for dispatched actions
    useEffect(() => {
      const unsubscribe = store.subscribe(() => {
        const state = store.getState();

        const connectivityStatus = state.connectivityStatus;
        if (connectivityStatus === "connected") {
          const lastActivity = state.activities?.[state.activities.length - 1];
          if (
            lastActivity?.type === "message" &&
            lastActivity.from?.role === "bot"
          ) {
            // Check for conversation ID changes
            const currentConversationId = lastActivity.conversation?.id;
            if (
              currentConversationId &&
              currentConversationId !== lastConversationId
            ) {
              setLastConversationId(currentConversationId);
              if (onConversationIdUpdate) {
                onConversationIdUpdate(currentConversationId);
                //setShouldSendHello(true);
              }
            }
            const activityMessage = lastActivity.text;
            if (activityMessage && onAgentMessageUpdate) {
              onAgentMessageUpdate(activityMessage);

            }
          }
        }
      });

      return () => {
        unsubscribe();
      };
    }, [
      store,
      onAgentMessageUpdate,
      onConversationIdUpdate,
      lastConversationId,
    ]);

    // Separate effect for sending messages (if needed)
    // const [shouldSendHello, setShouldSendHello] = useState(false);

    // useEffect(() => {
    //   if (shouldSendHello) {
    //     store.dispatch({
    //       type: "WEB_CHAT/SEND_MESSAGE",
    //       payload: { text: "Hello" },
    //     });
    //     setShouldSendHello(false);
    //   }
    // }, [shouldSendHello, store]);

    const chatWidth = typeof width === "number" ? `${width}px` : width;
    const chatHeight = typeof height === "number" ? `${height}px` : height;

    const isConfigured =
      appClientId &&
      tenantId &&
      (directConnectUrl || (environmentId && agentIdentifier));
    useEffect(() => {
      if (!isConfigured) return;

      let cancelled = false;

      const initializeConnection = async (): Promise<void> => {
        try {
          const token = await acquireToken({
            appClientId,
            tenantId,
            currentUserLogin,
            redirectUri: baseUrl,
          });

          if (!token) {
            setError("Unable to acquire token.");
            return;
          }

          const settings = new ConnectionSettings({
            appClientId,
            tenantId,
            environmentId: environmentId || "",
            agentIdentifier: agentIdentifier || "",
            directConnectUrl: directConnectUrl || "",
          });

          const clientInstance = new CopilotStudioClient(settings, token);
          const webchatSettings = { showTyping };

          if (!cancelled) {
            const connectionInstance =
              await CopilotStudioWebChat.createConnection(
                clientInstance,
                webchatSettings
              );

            try {
              setConnection(connectionInstance);
            } catch (error) {
              console.error("Error setting connection:", error);
              setError("Failed to establish connection");
            }

            //setConnection(
            //
            //);
          }
        } catch (e) {
          if (!cancelled) {
            console.error("Chat initialization error:", e);
            setError(
              e instanceof Error ? e.message : "Unknown error initializing chat"
            );
          }
        }
      };

      initializeConnection();

      return () => {
        cancelled = true;
      };
    }, [
      isConfigured,
      appClientId,
      tenantId,
      environmentId,
      agentIdentifier,
      directConnectUrl,
      showTyping,
      disableFileUploadButton,
      currentUserLogin,
      baseUrl,
      styleOptions,
      width,
      height,
    ]);

    React.useImperativeHandle(
      ref,
      () => ({
        sendMessage: (text: string) => {
          store.dispatch({
            type: "WEB_CHAT/SEND_MESSAGE",
            payload: { text },
          });
        },
        sendEvent: (name: string, value?: unknown) => {
          store.dispatch({
            type: "WEB_CHAT/SEND_EVENT",
            payload: { name, value },
          });
        },
        updateAgentMessage: (message: string) => {
          if (onAgentMessageUpdate) {
            onAgentMessageUpdate(message);
          }
        },
      }),
      [store, onAgentMessageUpdate, onConversationIdUpdate, onNewConversation]
    );

    if (!isConfigured) {
      return (
        <ChatMessage type="warning">
          Configure appClientId, tenantId, and either directConnectUrl or
          (environmentId and agentIdentifier) in the manifest properties.
        </ChatMessage>
      );
    }

    if (error) {
      return <ChatMessage type="error">Error: {error}</ChatMessage>;
    }

    if (!connection) {
      return <ChatMessage>Connecting to Copilot Studio...</ChatMessage>;
    }

    if (styleOptions && styleOptions != "val") {
      return (
        <div
          style={{
            height: chatHeight,
            width: chatWidth,
            gap: "8px",
            display: "flex",
            flexDirection: "column",
            padding: "16px 24px",
          }}
        >
            {!disableHeader && (
            <div
              style={{
              height: "50px",
              width: "88%",
              background: "white",
              borderRadius: "12px",
              padding: "16px 24px",
              boxShadow: "0 2px 8px rgba(0, 0, 0, 0.06)",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              }}
            >
              <h3
              style={{
                margin: 0,
                fontSize: 16,
                fontWeight: 600,
                color: "#323130",
              }}
              >
              {agentTitle}
              </h3>
              <div
              style={{
                display: "flex",
                gap: "8px",
                alignItems: "center",
              }}
              >
              <button
                style={{
                background: "none",
                border: "1px solid #605e5c",
                cursor: "pointer",
                padding: "8px",
                display: "flex",
                alignItems: "center",
                justifyContent: "left",
                color: "#605e5c",
                borderRadius: "4px",
                transition: "all 0.2s ease",
                }}
                aria-label="Start new conversation"
                onClick={onNewConversation}
                title="Start new conversation"
              >
                <span>
                <NewChatIcon />
                </span>
                &nbsp;New chat
              </button>
              </div>
            </div>
            )}
          <div style={{ height:  disableHeader ? chatHeight : `calc(${chatHeight} - 120px)`, width: "95%" }}>
            <ReactWebChat
              directLine={connection}
              store={store}
              styleOptions={JSON.parse(styleOptions)}
            ></ReactWebChat>
          </div>
        </div>
      );
    } else {
      return (
        <div
          id="chatContainer"
          style={{
            height: chatHeight,
            width: chatWidth,
            gap: "8px",
            display: "flex",
            flexDirection: "column",
          }}
        >
            {!disableHeader && (
            <div
              style={{
              height: "50px",
              width: chatWidth,
              background: "white",
              borderRadius: "12px",
              padding: "16px 24px",
              boxShadow: "0 2px 8px rgba(0, 0, 0, 0.06)",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              }}
            >
              <h3
              style={{
                margin: 0,
                fontSize: 16,
                fontWeight: 600,
                color: "#323130",
              }}
              >
              {agentTitle}
              </h3>
              <div
              style={{
                display: "flex",
                gap: "8px",
                alignItems: "center",
              }}
              >
              <button
                style={{
                background: "none",
                border: "1px solid #605e5c",
                cursor: "pointer",
                padding: "8px",
                display: "flex",
                alignItems: "center",
                justifyContent: "left",
                color: "#605e5c",
                borderRadius: "4px",
                transition: "all 0.2s ease",
                }}
                aria-label="Start new conversation"
                onClick={onNewConversation}
                title="Start new conversation"
              >
                <span>
                <NewChatIcon />
                </span>
                &nbsp;New chat
              </button>
              </div>
            </div>
            )}
            <div
            style={{
              height: disableHeader ? chatHeight : `calc(${chatHeight} - 60px)`,
              width: chatWidth,
              minHeight: "400px",
              minWidth: "200px",
            }}
            >
            <FluentThemeProvider fontSize="14px">
              <Composer
              directLine={connection}
              store={store}
              styleOptions={{
                rootHeight: "100%",
                rootWidth: "100%",
                disableFileUpload: disableFileUploadButton
              }}
              >
              <BasicWebChat/>
              </Composer>
            </FluentThemeProvider>
            </div>
        </div>
      );
    }
  }
);

interface ChatMessageProps {
  children: React.ReactNode;
  type?: "info" | "warning" | "error";
}

const ChatMessage: React.FC<ChatMessageProps> = ({
  children,
  type = "info",
}) => {
  const styles = {
    info: { padding: 16 },
    warning: { padding: 16, color: "#8a6d3b", backgroundColor: "#fcf8e3" },
    error: { padding: 16, color: "#a94442", backgroundColor: "#f2dede" },
  };

  return <div style={styles[type]}>{children}</div>;
};

export default Chat;

const NewChatIcon: React.FC = () => (
  <svg
    width="20"
    height="20"
    viewBox="0 0 20 20"
    fill="none"
    xmlns="http://www.w3.org/2000/svg"
  >
    <path
      d="M10.5 2C15.1944 2 19 5.80558 19 10.5C19 15.1944 15.1944 19 10.5 19C8.76472 19 7.11922 18.4543 5.75373 17.4816L2.49213 18.5078C2.08002 18.6317 1.65087 18.4024 1.52705 17.9903C1.48179 17.8421 1.48179 17.6842 1.52705 17.536L2.5527 14.2752C1.57831 12.9086 1.0317 11.2612 1.0317 9.52322C1.08606 5.05327 4.99567 1.31044 9.46027 1.03543C9.80475 1.01197 10.1513 1 10.5 1V2ZM10.5 3C6.35786 3 3 6.35786 3 10.5C3 11.8905 3.38968 13.1911 4.06306 14.2901C4.14846 14.4308 4.20128 14.5899 4.21725 14.7554C4.23321 14.9209 4.21188 15.0886 4.15513 15.2446L3.45116 17.3484L5.55498 16.6445C5.86607 16.5406 6.20598 16.5808 6.48684 16.7547C7.58156 17.4263 8.87622 17.8142 10.2598 17.8142H10.5C14.6421 17.8142 18 14.4563 18 10.3142V10.186C17.8591 6.33426 14.9119 3.21269 11.0732 3.00673C10.8831 2.99556 10.692 2.99 10.5 2.99V3ZM10.5 5.5C10.7761 5.5 11 5.72386 11 6V9.5H14.5C14.7761 9.5 15 9.72386 15 10C15 10.2761 14.7761 10.5 14.5 10.5H11V14C11 14.2761 10.7761 14.5 10.5 14.5C10.2239 14.5 10 14.2761 10 14V10.5H6.5C6.22386 10.5 6 10.2761 6 10C6 9.72386 6.22386 9.5 6.5 9.5H10V6C10 5.72386 10.2239 5.5 10.5 5.5Z"
      fill="currentColor"
    />
  </svg>
);
