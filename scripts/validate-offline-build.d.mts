export type OfflineBuildIssue = {
  match: RegExpMatchArray;
  name: string;
  pattern: RegExp;
};

export function neuterAnchorHrefs(html: string): string;
export function findOfflineBuildIssues(html: string): OfflineBuildIssue[];
