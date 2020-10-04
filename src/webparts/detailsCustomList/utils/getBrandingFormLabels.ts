import { getBrandingFormProperty } from "./getLookUpFields";

let tempObjLabels = {
  Name: "",
  Email: "",
  Department: "",
  DisclaimerParagraph: "",
  DisclaimerBullets: "",
  NominatedContentOwner: "",
  SubmitSuccess: "",
  SubmitFailed: "",
  BrandingRequired: "",
  SavedMessage: ""
};

export const formLabels = () => {
  getBrandingFormProperty().then(obj => {
    console.log("_getBrandingFormLabels obj", obj);
    obj.map(res => {
      if (res.Property === "Name") {
        tempObjLabels.Name = res.Text;
      }
      if (res.Property === "Email") {
        tempObjLabels.Email = res.Text;
      }
      if (res.Property === "Department") {
        tempObjLabels.Department = res.Text;
      }
      if (res.Property === "DisclaimerParagraph") {
        tempObjLabels.DisclaimerParagraph = res.Text;
      }
      if (res.Property === "DisclaimerBullets") {
        tempObjLabels.DisclaimerBullets = res.Text;
      }
      if (res.Property === "NominatedContentOwner") {
        tempObjLabels.NominatedContentOwner = res.Text;
      }
      if (res.Property === "SubmitSuccess") {
        tempObjLabels.SubmitSuccess = res.Text;
      }
      if (res.Property === "SubmitFailed") {
        tempObjLabels.SubmitFailed = res.Text;
      }
      if (res.Property === "BrandingRequired") {
        tempObjLabels.BrandingRequired = res.Text;
      }
      if (res.Property === "SavedMessage") {
        tempObjLabels.SavedMessage = res.Text;
      }
    });
  });
  return tempObjLabels;
};
