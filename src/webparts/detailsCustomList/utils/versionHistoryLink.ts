import SharePointService from "../services/SharePointService";

export const versionHistoryLink = (listId: string, itemId: string): string => {
  const absoluteUrl = SharePointService.context.pageContext.web.absoluteUrl;
  return `${absoluteUrl}/_layouts/15/Versions.aspx?List=${listId}&ID=${itemId}&IsDlg=1`;
};
