import { IFeedbackField } from "./IFeedbackField";

export interface IFeedbackForm {
  activateFeedbackForm: boolean;
  feedbackListName: string;
  feedbackFields: IFeedbackField[];
}
