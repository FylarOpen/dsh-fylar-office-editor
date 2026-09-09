/** Turn-scoped Office outputs and their assistant-tail file cards. */
import type { ConversationNodeDefinition } from '@deepseek-ai/dsh-client-ui-conversation/client';
import type { TurnTailOwnerProps } from '@deepseek-ai/dsh-client-ui-chat/client';
import type { InjectFace, PropsLocale, PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots';
import type { OfficeArtifactTarget, OfficeSessionSnapshot } from './office-artifact-controller.ts';
import type { ObservableSnapshot, OfficeWorkspaceActions, OfficeWorkspaceSnapshot } from './office-workspace.ts';
import { type OfficeDeliverablesData } from './office-deliverables-model.ts';
import { OFFICE_LOCALE_NS } from './locales.ts';
export interface OfficeDeliverablesTurnData extends OfficeDeliverablesData {
}
/** Conversation Definition kind and the Turn-data key it exclusively owns. */
export declare const OFFICE_DELIVERABLES_KIND = "fylar-office-deliverables";
declare module '@deepseek-ai/dsh-client-ui-conversation/client' {
    interface ConversationTurnDataMap {
        /** Office files created or published by a dsh-fylar-office-editor Tool in this turn. */
        'fylar-office-deliverables': OfficeDeliverablesTurnData;
    }
}
interface OfficeDeliverablesState extends OfficeDeliverablesTurnData {
    readonly turn: number;
    readonly calls: ReadonlyMap<string, string>;
}
/** Accumulate successful Office creation calls into their engine-owned Turn. */
export declare const officeDeliverablesDefinition: ConversationNodeDefinition<OfficeDeliverablesState>;
export declare function selectOfficeDeliverables(owner: TurnTailOwnerProps): readonly string[] | null;
export interface OfficeDeliverablesInjected extends Pick<OfficeWorkspaceActions, 'attach' | 'detach' | 'setReadOnly' | 'exportDocument' | 'exportPdf' | 'clearIssue' | 'markPotentialEdits'> {
    hooks: {
        workspace: ObservableSnapshot<OfficeWorkspaceSnapshot>;
        officeSession: ObservableSnapshot<OfficeSessionSnapshot>;
    };
    openDetails: (target: OfficeArtifactTarget) => void;
    openInlineArtifact: (target: OfficeArtifactTarget) => void;
    collapseInline: () => void;
    closeDocument: () => void;
    retrySession: () => Promise<void>;
}
type OfficeDeliverablesProps = PropsRuntime<'conversation.chat.turnTail'> & {
    matched: readonly string[];
} & PropsLocale<typeof OFFICE_LOCALE_NS> & InjectFace<OfficeDeliverablesInjected>;
/** Render successful files after the closing assistant response. */
export declare function OfficeDeliverables({ matched, useSessions, sessionId, useWorkspace, useOfficeSession, openDetails, openInlineArtifact, collapseInline, closeDocument, retrySession, attach, detach, setReadOnly, exportDocument, exportPdf, clearIssue, markPotentialEdits, t, }: OfficeDeliverablesProps): import("react/jsx-runtime").JSX.Element;
export {};
//# sourceMappingURL=OfficeDeliverables.d.ts.map