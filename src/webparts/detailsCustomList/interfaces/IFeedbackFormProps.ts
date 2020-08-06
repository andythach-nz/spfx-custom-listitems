import { IFeedbackForm } from "./IFeedbackForm";
import { ISelectedItem } from "./ISelectedItem";

export interface IFeedbackFormProps {
  feedbackFormSettings: IFeedbackForm;
  docId: string;
  isOpen: boolean;
  onCloseForm: () => void;
  stream: string;
  selectedItems: ISelectedItem[];
}
