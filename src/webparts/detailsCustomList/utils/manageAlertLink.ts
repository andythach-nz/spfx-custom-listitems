import SharePointService from "../services/SharePointService";

export const manageAlertLink = (): string => {
  const absoluteUrl = SharePointService.context.pageContext.web.absoluteUrl;
  return `${absoluteUrl}/_layouts/15/mysubs.aspx`;
};
