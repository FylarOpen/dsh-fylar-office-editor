/** Browser plugin linking Agent-created Office files to Chat, Office, and the right Sidebar. */
import type { Context as ClientContext } from '@deepseek-ai/cordis';
export declare const OFFICE_RIGHTBAR_ID = "@fylar/dsh-fylar-office-editor";
export declare const OFFICE_RIGHTBAR_KIND = "fylar-office";
/** Browser services required for Tool status, Chat delivery, Office, and the right Sidebar. */
export declare const inject: string[];
/** Register one Office session across Chat, the Office view, and the right Sidebar. */
export declare function apply(ctx: ClientContext): void;
//# sourceMappingURL=index.d.ts.map