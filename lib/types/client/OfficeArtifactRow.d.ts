/** Compact progress row for Office creation and existing-file publication. */
import type { ToolCallViewProps } from '@deepseek-ai/dsh-client-ui-tool/client';
import type { PropsLocale } from '@deepseek-ai/dsh-client-ui-slots';
import { OFFICE_LOCALE_NS } from './locales.ts';
export type OfficeArtifactRowProps = ToolCallViewProps & PropsLocale<typeof OFFICE_LOCALE_NS>;
interface ToolRowModel {
    readonly path: string | undefined;
    readonly state: 'running' | 'ready' | 'error' | 'stopped';
}
export declare function officeArtifactRowModel(block: ToolCallViewProps['block']): ToolRowModel;
/** Render generation progress only; the assistant-tail card is the file entry. */
export declare function OfficeArtifactRow({ block, toolName, t }: OfficeArtifactRowProps): import("react/jsx-runtime").JSX.Element;
export {};
//# sourceMappingURL=OfficeArtifactRow.d.ts.map