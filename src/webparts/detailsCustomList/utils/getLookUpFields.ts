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

export const getBrandingReasons = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems("LOOKUPBrandingReason");
  } catch (error) {
    throw error;
  }
};

export const getBrandingOptions = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems("Branding%20Configuration");
  } catch (error) {
    throw error;
  }
};

export const getSavedDocuments = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems("SavedBrandingDocument");
  } catch (error) {
    throw error;
  }
};

export const getSaveListItem = async (
  getLibName: string,
  getItemId: number,
  expand: string[]
): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItem(
      getLibName,
      getItemId,
      expand
    );
  } catch (error) {
    throw error;
  }
};

export const getBrandingFormProperty = async (): Promise<any> => {
  try {
    return await SharePointService.pnp_getListItems(
      "LOOKUPBrandingFormProperty"
    );
  } catch (error) {
    throw error;
  }
};
