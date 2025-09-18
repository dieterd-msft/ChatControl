import * as React from "react";
import { useEffect, useState } from "react";
import { Components, createStore } from "botframework-webchat";
import { FluentThemeProvider } from "botframework-webchat-fluent-theme";
import {
  ConnectionSettings,
  CopilotStudioClient,
  CopilotStudioWebChat,
  CopilotStudioWebChatConnection,
} from "@microsoft/agents-copilotstudio-client";
import { acquireToken } from "./acquireToken";

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
  currentUserLogin?: string;
  baseUrl?: string;
  styleOptions?: string;
  width?: string |number;
  height?: string | number;
  onAgentMessageUpdate?: (message: string) => void;
  onNewConversation?: () => void;
}

export interface ChatRef {
  sendMessage: (text: string) => void;
  getConnectionState: () => string | undefined;
  updateAgentMessage: (message: string) => void;
}

const Chat = React.forwardRef<ChatRef, IChatProps>(
  (
    {
      chatKey,
      agentTitle,
      appClientId,
      tenantId,
      environmentId,
      agentIdentifier,
      directConnectUrl,
      showTyping = true,
      currentUserLogin,
      baseUrl,
      width,
      height,
      onAgentMessageUpdate,
      onNewConversation,
    },
    ref
  ) => {
    const [connection, setConnection] =
      useState<CopilotStudioWebChatConnection | null>(null);
    const [error, setError] = useState<string>();
    const [store] = useState(() => createStore());

    // Set up store listener for dispatched actions
    useEffect(() => {
      const unsubscribe = store.subscribe(() => {
        const state = store.getState();

        // console.log('Store action dispatched. Current state:', {
        //   activities: state.activities?.length || 0
        // });
        // state.activities?.forEach((activity:object) => {
        //   console.log('New activity:', activity);
        // });
        const lastActivity = state.activities?.[state.activities.length - 1];
        if (
          lastActivity?.type === "message" &&
          lastActivity.from?.role === "bot"
        ) {
          const activityMessage = lastActivity.text;
          if (activityMessage && onAgentMessageUpdate) {
            onAgentMessageUpdate(activityMessage);
          }
        }
      });

      return () => {
        unsubscribe();
      };
    }, [store, onAgentMessageUpdate, onNewConversation]);

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
            setConnection(
              CopilotStudioWebChat.createConnection(
                clientInstance,
                webchatSettings
              )
            );
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
      currentUserLogin,
      baseUrl,
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
        getConnectionState: () => {
          const state = store.getState();
          return state.connectionStatus;
        },
        updateAgentMessage: (message: string) => {
          if (onAgentMessageUpdate) {
            onAgentMessageUpdate(message);
          }
        },
      }),
      [store, onAgentMessageUpdate, onNewConversation]
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

    return (
      <div style={{ height, width }}>
        <div>
              <h3>{agentTitle}</h3>
              <div>
                <button
                  aria-label="Start new conversation"
                  onClick={onNewConversation}
                  title="Start new conversation"
                >
                  <NewChatIcon />
                </button>
              </div>
            </div>
            <div>
      <FluentThemeProvider>
        <Composer directLine={connection} store={store}>
          <BasicWebChat />
        </Composer>
      </FluentThemeProvider></div>
      </div>
    );
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
  <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M10.5 2C15.1944 2 19 5.80558 19 10.5C19 15.1944 15.1944 19 10.5 19C8.76472 19 7.11922 18.4543 5.75373 17.4816L2.49213 18.5078C2.08002 18.6317 1.65087 18.4024 1.52705 17.9903C1.48179 17.8421 1.48179 17.6842 1.52705 17.536L2.5527 14.2752C1.57831 12.9086 1.0317 11.2612 1.0317 9.52322C1.08606 5.05327 4.99567 1.31044 9.46027 1.03543C9.80475 1.01197 10.1513 1 10.5 1V2ZM10.5 3C6.35786 3 3 6.35786 3 10.5C3 11.8905 3.38968 13.1911 4.06306 14.2901C4.14846 14.4308 4.20128 14.5899 4.21725 14.7554C4.23321 14.9209 4.21188 15.0886 4.15513 15.2446L3.45116 17.3484L5.55498 16.6445C5.86607 16.5406 6.20598 16.5808 6.48684 16.7547C7.58156 17.4263 8.87622 17.8142 10.2598 17.8142H10.5C14.6421 17.8142 18 14.4563 18 10.3142V10.186C17.8591 6.33426 14.9119 3.21269 11.0732 3.00673C10.8831 2.99556 10.692 2.99 10.5 2.99V3ZM10.5 5.5C10.7761 5.5 11 5.72386 11 6V9.5H14.5C14.7761 9.5 15 9.72386 15 10C15 10.2761 14.7761 10.5 14.5 10.5H11V14C11 14.2761 10.7761 14.5 10.5 14.5C10.2239 14.5 10 14.2761 10 14V10.5H6.5C6.22386 10.5 6 10.2761 6 10C6 9.72386 6.22386 9.5 6.5 9.5H10V6C10 5.72386 10.2239 5.5 10.5 5.5Z" fill="currentColor"/>
  </svg>
);
