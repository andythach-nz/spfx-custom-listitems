import SharePointService from "../services/SharePointService";

export const getFeedbackTypes = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems("LOOKUPFeedbackType");
  } catch (error) {
    throw error;
  }
};

export const getFeedbackCategoriesPage = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems("LOOKUPFeedbackCategory");
  } catch (error) {
    throw error;
  }
};

export const getFeedbackCategoriesDocument = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems(
      "LOOKUPFeedbackCategoryDocument"
    );
  } catch (error) {
    throw error;
  }
};

export const getFeedbackAreas = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems("LOOKUPFeedbackArea");
  } catch (error) {
    throw error;
  }
};
