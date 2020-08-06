import * as React from "react";
import {
  PrimaryButton,
  TextField,
  Stack,
  IStackTokens,
  Panel,
  PanelType,
  Label,
  Separator,
  Image,
  Dropdown,
  IDropdownOption,
  Link
} from "office-ui-fabric-react";
import SharePointService from "../services/SharePointService";
import { IFeedbackFormProps } from "../interfaces/IFeedbackFormProps";
import {
  FilePicker,
  IFilePickerResult
} from "@pnp/spfx-controls-react/lib/FilePicker";
import { fieldColumnMapper } from "../mappers/FeedbackFormMapper";
import { IFeedbackFormState } from "../interfaces/IFeedbackFormState";
import { AttachmentFileInfo } from "@pnp/sp";
import {
  getFeedbackTypes,
  getFeedbackCategoriesPage,
  getFeedbackCategoriesDocument,
  getFeedbackAreas
} from "../utils/getLookUpFields";
const formImage: string = require("../images/FormResource.png");

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 2
};

export const FeedbackForm = ({
  isOpen,
  onCloseForm,
  feedbackFormSettings,
  docId,
  stream,
  selectedItems
}: IFeedbackFormProps): JSX.Element => {
  const [isDisable, setIsDisable] = React.useState<boolean>(false);
  const [isCategoryTypeDisable, setIsCategoryTypeDisable] = React.useState<
    boolean
  >(false);
  const [listOfAttachments, setListOfAttachments] = React.useState<
    AttachmentFileInfo[]
  >([]);
  const [feedback, setFeedback] = React.useState<string | undefined>(undefined);
  const [feedbackFormValues, setfeedbackFormValues] = React.useState<
    IFeedbackFormState
  >(fieldColumnMapper(feedbackFormSettings.feedbackFields));

  const [feedbackTypeOptions, setFeedbackTypeOptions] = React.useState<
    IDropdownOption[]
  >();
  const [feedbackCategoryForPage, setFeedbackCategoryForPage] = React.useState<
    IDropdownOption[]
  >();
  const [
    feedbackCategoryForDocuments,
    setFeedbackCategoryForDocuments
  ] = React.useState<IDropdownOption[]>();
  const [feedbackCategories, setFeedbackCategories] = React.useState<
    IDropdownOption[]
  >();
  const [feedbackAreas, setFeedbackAreas] = React.useState<IDropdownOption[]>();
  const [feedbackSelectedAreas, setFeedbackSelectedAreas] = React.useState<
    string
  >();

  let fileInfos: AttachmentFileInfo[] = listOfAttachments;

  React.useEffect(() => {
    let documentNamesArr = [],
      tempDocuLinksArr = [];
    selectedItems.map(item => {
      documentNamesArr.push(item.selectedItemName);
      tempDocuLinksArr.push(item.serverRelativeUrl);
    });
    let documentNames = documentNamesArr.join(";");
    let tempDocuLinks = tempDocuLinksArr.join(";");

    const loginName =
      "i:0#.f|membership|" +
      SharePointService.context.pageContext.user.loginName;
    SharePointService.pnp_getUserProfileProperty(loginName, "Department").then(
      res => {
        setfeedbackFormValues({
          DocumentName: documentNames,
          DocumentLinks: tempDocuLinks,
          Title: "View full submission",
          Email: SharePointService.context.pageContext.user.email,
          Department: res,
          FeedbackType: "",
          FeedbackCategory: "",
          Feedback: "",
          Stream: stream,
          PageURL: window.location.href,
          AssignedToEmail: SharePointService.context.pageContext.user.loginName,
          feedbackTypeSelectionKeys: [],
          feedbackCategorySelectionKeys: []
        });
        setIsDisable(true);
        setIsCategoryTypeDisable(true);
      }
    );
  }, []);

  React.useEffect(() => {
    getFeedbackTypes().then(r => {
      const typeResults = r.map(item => ({
        key: item.Title,
        text: item.Title
      }));
      setFeedbackTypeOptions(typeResults);
    });
  }, []);

  React.useEffect(() => {
    getFeedbackCategoriesPage().then(res => {
      const categoryResults = res.map(item => ({
        key: item.Title,
        text: item.Title
      }));
      setFeedbackCategoryForPage(categoryResults);
    });

    getFeedbackCategoriesDocument().then(res => {
      const categoryDocumentResults = res.map(item => ({
        key: item.Title,
        text: item.Title
      }));
      setFeedbackCategoryForDocuments(categoryDocumentResults);
    });
  }, []);

  React.useEffect(() => {
    console.log("--change");
  }, [feedbackFormValues.feedbackTypeSelectionKeys]);

  React.useEffect(() => {
    getFeedbackAreas().then(res => {
      const areaResults = res.map(item => ({
        key: item.key,
        text: item.Title
      }));

      setFeedbackAreas(areaResults);

      areaResults.map(areaRes => {
        if (areaRes.key === stream) {
          setFeedbackSelectedAreas(areaRes.key);
        }
      });
    });
  }, []);

  const _handleFileAttachment = async (filePickerResult: IFilePickerResult) => {
    if (filePickerResult.fileAbsoluteUrl == null) {
      filePickerResult.downloadFileContent().then(async fileR => {
        let tempAttachmentsArray = fileInfos;
        tempAttachmentsArray.push({ name: fileR.name, content: fileR });
        fileInfos = tempAttachmentsArray;
        setfeedbackFormValues({
          ...feedbackFormValues,
          feedbackAttachments: fileInfos
        });
        console.log("fileInfos", fileInfos);
      });
    }
  };

  const _handleOnChange = (
    e: React.FormEvent<HTMLInputElement>,
    inputValue: IDropdownOption,
    index: any
  ) => {
    const currentId = e.target["id"] as string;
    if (currentId === "feedbackTypeSelectionKeys") {
      if (inputValue.text.includes("page")) {
        setIsDisable(false);
        setFeedbackCategories(feedbackCategoryForPage);
      } else {
        setIsDisable(true);
        setFeedbackCategories(feedbackCategoryForDocuments);
      }

      inputValue.text.includes("Select")
        ? setIsCategoryTypeDisable(true)
        : setIsCategoryTypeDisable(false);
      setfeedbackFormValues({
        ...feedbackFormValues,
        feedbackTypeSelectionKeys: inputValue
      });
    } else if (currentId === "feedbackCategorySelectionKeys") {
      let tempCategoryArray =
        feedbackFormValues["feedbackCategorySelectionKeys"];
      if (inputValue.selected) {
        tempCategoryArray.push(inputValue.key);
      } else {
        tempCategoryArray = feedbackFormValues[
          "feedbackCategorySelectionKeys"
        ].filter(t => t !== inputValue.key);
      }
      setfeedbackFormValues({
        ...feedbackFormValues,
        feedbackCategorySelectionKeys: tempCategoryArray
      });
    }
  };

  const _submitForm = async (e: any): Promise<void> => {
    e.preventDefault();
    const submitObj = feedbackFormValues;
    //const optionTypeValues = submitObj.feedbackTypeSelectionKeys.join(";");
    const optionCategoryValues = submitObj.feedbackCategorySelectionKeys.join(
      ";"
    );
    submitObj.FeedbackType = feedbackFormValues.feedbackTypeSelectionKeys.text;
    submitObj.FeedbackCategory = optionCategoryValues;
    delete submitObj.feedbackTypeSelectionKeys;
    delete submitObj.feedbackCategorySelectionKeys;
    delete submitObj.feedbackAttachments;
    try {
      const itemResult = await SharePointService.pnp_addItem(
        feedbackFormSettings.feedbackListName,
        submitObj
      );
      //handle attachments
      itemResult.item.attachmentFiles.addMultiple(listOfAttachments);

      onCloseForm();
    } catch (error) {
      onCloseForm();
      throw error;
    }
  };

  const _removeAttachments = (e: any) => {
    console.log("_removeAttachments called", e.target.value);
    let tempAttachments = feedbackFormValues["feedbackAttachments"];
    tempAttachments = feedbackFormValues["feedbackAttachments"].filter(
      item => item.name !== e.target.value
    );
    setListOfAttachments(tempAttachments);
    setfeedbackFormValues({
      ...feedbackFormValues,
      feedbackAttachments: tempAttachments
    });
  };

  const _onRenderDisplayFormFields = (): JSX.Element[] =>
    feedbackFormSettings.feedbackFields.map(row => {
      if (row.valueMarker === "displayName") {
        return (
          <div style={{ display: "inline-flex" }}>
            <Label style={{ fontWeight: 500, paddingRight: "5px" }}>
              Name:{" "}
            </Label>
            <Label>
              {SharePointService.context.pageContext.user.displayName}
            </Label>
          </div>
        );
      } else if (row.valueMarker === "email") {
        return (
          <div style={{ display: "inline-flex" }}>
            <Label style={{ fontWeight: 500, paddingRight: "5px" }}>
              Email:{" "}
            </Label>
            <Label>
              {" " + SharePointService.context.pageContext.user.email}
            </Label>
          </div>
        );
      } else if (row.valueMarker === "department") {
        return (
          <div style={{ display: "inline-flex" }}>
            <Label style={{ fontWeight: 500, paddingRight: "5px" }}>
              Department:{" "}
            </Label>
            <Label>
              {" " + feedbackFormValues.Department
                ? feedbackFormValues.Department
                : ""}
            </Label>
          </div>
        );
      } else if (row.valueMarker === "feedbackType") {
        return (
          <Dropdown
            placeholder="Select an option"
            label="I would like to submit feedback regarding:"
            id="feedbackTypeSelectionKeys"
            options={feedbackTypeOptions}
            onChange={_handleOnChange}
          />
        );
      }
    });

  const _onRenderFooterContent = () => {
    return (
      <Stack horizontal horizontalAlign="end">
        <PrimaryButton
          onClick={_submitForm}
          text="Submit"
          disabled={
            !feedbackFormValues.Feedback ||
            feedbackFormValues.feedbackTypeSelectionKeys.length === 0 ||
            feedbackFormValues.feedbackCategorySelectionKeys.length === 0 ||
            (!feedbackFormValues.PageURL && !isDisable)
          }
        />
      </Stack>
    );
  };

  const _onRenderHeader = (): JSX.Element => {
    return (
      <Stack verticalAlign="center">
        <Image
          src={formImage}
          width={300}
          height={110}
          styles={{ root: { margin: "0 auto" } }}
        />
        <Label style={{ fontSize: 13, textAlign: "center", margin: 10 }}>
          Hi {SharePointService.context.pageContext.user.displayName}, when you
          submit this form, the owner will be able to see your name and email
          address.
        </Label>
        <Separator />
      </Stack>
    );
  };

  const getFeedbackTypeSelectionKeys = React.useCallback(() => {
    const { feedbackTypeSelectionKeys } = feedbackFormValues;
    console.log("opp", feedbackFormSettings.feedbackFields);
    console.log(feedbackTypeSelectionKeys);
    return feedbackTypeSelectionKeys;
  }, [feedbackFormValues]);
  return (
    <Panel
      styles={{ headerText: { fontFamily: "Arial, Helvetica, sans-serif" } }}
      isOpen={isOpen}
      type={PanelType.custom}
      customWidth="520px"
      onDismiss={onCloseForm}
      closeButtonAriaLabel="Close"
      onRenderFooterContent={_onRenderFooterContent}
      onRenderHeader={_onRenderHeader}
    >
      <form onSubmit={_submitForm}>
        <Stack tokens={itemAlignmentsStackTokens}>
          <Label>
            <p>
              Thank you for taking the time to submit feedback. Your name, email
              address and feedback will be sent to the relevant owner for
              consideration.
            </p>
            <p>
              Note: For all IT related issues, Australian users contact 1300 333
              000, New Zealand Users contact 0800 156 666 and Spotless Users
              Contact - AU - 1300 333 000, NZ - 0800 487 768
            </p>
          </Label>

          {_onRenderDisplayFormFields()}

          <Dropdown
            isDisabled={isCategoryTypeDisable}
            placeholder="Select an option"
            multiSelect
            label="Feedback Category:"
            id="feedbackCategorySelectionKeys"
            options={feedbackCategories}
            onChange={_handleOnChange}
            selectedKeys={[...feedbackFormValues.feedbackCategorySelectionKeys]}
          />

          <TextField
            label="Document(s) selected:"
            value={feedbackFormValues.DocumentName}
          />

          <TextField
            label="Page URL:"
            disabled={isDisable}
            required={!isDisable}
            value={feedbackFormValues.PageURL}
            onChange={(e: any, newValue?: string) => {
              setfeedbackFormValues({
                ...feedbackFormValues,
                PageURL: newValue
              });
            }}
          />

          <Dropdown
            placeholder="Select an option"
            label="The feedback is for the following areas:"
            id="feedbackStreamOptions"
            options={feedbackAreas}
            onChange={_handleOnChange}
            selectedKey={feedbackSelectedAreas}
          />

          <TextField
            multiline
            label="Please provide more information:"
            rows={4}
            onChange={(e: any, newValue?: string) =>
              setfeedbackFormValues({
                ...feedbackFormValues,
                Feedback: newValue
              })
            }
            value={feedback}
            placeholder="Your feedback..."
          />

          <FilePicker
            bingAPIKey="ArMip2E-OGmxgNXdjqFvjKPsIkUu8tfWlsaIROS3vWkl26KCQpVdVLD2ua63-bOr"
            hideWebSearchTab={true}
            accepts={[
              ".gif",
              ".jpg",
              ".jpeg",
              ".bmp",
              ".dib",
              ".tif",
              ".tiff",
              ".ico",
              ".png",
              ".jxr",
              ".svg"
            ]}
            buttonIcon="FileImage"
            buttonLabel="Upload a File"
            onSave={(file?: IFilePickerResult) => {
              _handleFileAttachment(file);
            }}
            onChanged={(file?: IFilePickerResult) => {
              _handleFileAttachment(file);
            }}
            context={SharePointService.context}
            label="Please attach file related to this feedback."
          />

          {feedbackFormValues.feedbackAttachments && (
            <div>
              <label style={{ fontWeight: "bold" }}>Attachment(s):</label>
              <ul>
                {feedbackFormValues.feedbackAttachments.map(file => {
                  return (
                    <li
                      style={{
                        listStyle: "none"
                      }}
                    >
                      {file.name}{" "}
                      <Link value={file.name} onClick={_removeAttachments}>
                        remove
                      </Link>
                    </li>
                  );
                })}
              </ul>
            </div>
          )}
        </Stack>
      </form>
    </Panel>
  );
};
