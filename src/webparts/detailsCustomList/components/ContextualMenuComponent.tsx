import * as React from "react";
import {
  ActionButton,
  ContextualMenuItemType,
  DirectionalHint,
  Callout
} from "office-ui-fabric-react";
import { AlertMeForm } from "./AlertMeForm";
import { SPFieldsContext } from "../contexts/SPFieldsContext";
import { SPItemsContext } from "../contexts/SPItemsContext";
import { FeedbackContext } from "../contexts/FeedbackContext";
import { FeedbackForm } from "./FeedbackForm";
import { alertMeLink } from "../utils/alertMeLink";
import { getOpentInLink } from "../utils/openInLink";
import { copyLink } from "../utils/copyLink";
import { ShareLinkForm } from "./ShareLinkForm";
import { dowloadSingleFile } from "../utils/dowloadSingleFile";
import { getZippedFiles } from "../utils/getZippedFiles";
import { IContextualMenuComponentProps } from "../interfaces/IContextualMenuComponentProps";
import { VersionHistoryForm } from "../components/VersionHistoryForm";
import { versionHistoryLink } from "../utils/versionHistoryLink";
import { manageAlertLink } from "../utils/manageAlertLink";
import { ISelectedItem } from "../interfaces/ISelectedItem";
import { DocumentBrandForm } from "./DocumentBrandForm";
import { getSavedDocuments } from "../utils/getLookUpFields";
import SharePointService from "../services/SharePointService";

const checkMultiForAspx = (selectedItems: ISelectedItem[]) => {
  let results = "";
  if (selectedItems.length === 1) {
    for (let i = 0; i < selectedItems.length; i++) {
      if (
        selectedItems[i].selectedItemExt === "aspx" ||
        selectedItems[i].selectedItemExt === "url"
      ) {
        results = "none";
        break;
      } else {
        results = "inline-block";
      }
    }
  } else if (selectedItems.length > 1) {
    for (let i = 0; i < selectedItems.length; i++) {
      if (
        selectedItems[i].selectedItemExt !== "aspx" ||
        selectedItems[i].selectedItemExt !== "url"
      ) {
        results = "inline-block";
        break;
      } else {
        results = "none";
      }
    }
  } else {
    results = "none";
  }
  return results;
};

export const ContextualMenuComponent: React.FC<IContextualMenuComponentProps> = React.memo(
  ({ selectedItemId, docId, stream }): JSX.Element => {
    const contextualMneuDialogRef = React.useRef();
    const [isCopyLinkDialog, setIsCopyLinkDialog] = React.useState(false);
    const [isShareLinkDialog, setIsShareLinkDialog] = React.useState(false);
    const [isAlerMeDialog, setAlerMeDialog] = React.useState<boolean>(false);
    const [isFeedbackForm, setFeedbackForm] = React.useState<boolean>(false);
    const [isVersionHistoryForm, setVersionHistoryForm] = React.useState(false);
    const [isDocumentBrandForm, setDocumentBrandForm] = React.useState(false);
    const [isDocumentBrandSaved, setDocumentBrandSaved] = React.useState(false);
    const { selectedListId, selectedListInternalName } = React.useContext(
      SPFieldsContext
    );
    const {
      selectedItems,
      setSelectedItems,
      setClearSelection
    } = React.useContext(SPItemsContext);
    const { feedbackForm } = React.useContext(FeedbackContext);

    React.useEffect(() => {
      getSavedDocuments().then(obj => {
        obj.map(i => {
          switch (
            i.Email === SharePointService.context.pageContext.user.email
          ) {
            case true:
              setDocumentBrandSaved(true);
              break;
            case false:
              setDocumentBrandSaved(false);
              break;
            default:
              setDocumentBrandSaved(false);
          }
        });
      });
      console.log("isDocumentBrandSaved", isDocumentBrandSaved);
    }, []);

    const _handleOnClose = () => {
      setDocumentBrandForm(false);
      setDocumentBrandSaved(false);
    };

    return (
      <div className="calloutArea">
        <ActionButton
          persistMenu={true}
          menuProps={{
            directionalHint: DirectionalHint.bottomCenter,
            shouldFocusOnMount: true,
            shouldFocusOnContainer: true,
            items: [
              {
                key: "openInApp",
                subMenuProps: {
                  items: [
                    {
                      key: "openInBrowser",
                      text: "Open in browser",
                      title: "Open in browser",
                      href:
                        selectedItems.length > 0 &&
                        selectedItems[0].selectedItemUrlOpenInBrowser,
                      target: "_blank",
                      ["data-interception"]: "off"
                    },
                    {
                      key: "openInApp",
                      text: "Open in app",
                      title: "Open in app",
                      href:
                        selectedItems.length > 0 &&
                        getOpentInLink(
                          selectedItems[0].selectedItemExt,
                          selectedListInternalName,
                          selectedItems[0].selectedItemName
                        )
                    }
                  ]
                },
                text: "Open",
                style: {
                  display:
                    selectedItems.length === 1 &&
                    selectedItems[0].selectedItemExt !== "aspx"
                      ? "inline-block"
                      : "none"
                }
              },
              {
                key: "divider_1",
                itemType: ContextualMenuItemType.Divider
              },
              // {
              //   key: "share",
              //   text: "Share",
              //   onClick: () => setIsShareLinkDialog(true),
              //   style: {
              //     display: selectedItems.length === 1 ? "inline-block" : "none"
              //   }
              // },
              {
                key: "copyLink",
                text: "Copy link",
                onClick: () => setIsCopyLinkDialog(true),
                style: {
                  display: selectedItems.length === 1 ? "inline-block" : "none"
                }
              },
              {
                key: "download",
                text: "Download",
                href:
                  selectedItems.length === 1 &&
                  dowloadSingleFile(selectedItems[0]),
                onClick:
                  selectedItems.length > 1
                    ? async () => await getZippedFiles(selectedItems)
                    : () => null,
                style: {
                  display:
                    selectedItems.length === 1 &&
                    selectedItems[0].selectedItemExt === "aspx"
                      ? "none"
                      : "inline-block"
                }
              },
              {
                key: "alertMe",
                text: "Alert Me",
                onClick: () => setAlerMeDialog(true),
                style: {
                  display:
                    selectedItems.length === 1 &&
                    selectedItems[0].selectedItemExt === "aspx"
                      ? "none"
                      : "inline-block"
                }
              },
              {
                key: "manageAlerts",
                text: "Manage My Alerts",
                href: manageAlertLink(),
                target: "_blank",
                ["data-interception"]: "off"
              },

              {
                key: "feedback",
                text: "Feedback",
                onClick: () => setFeedbackForm(true),
                style: {
                  display:
                    feedbackForm &&
                    selectedItems.length === 1 &&
                    selectedItems[0].selectedItemExt !== "aspx"
                      ? "inline-block"
                      : "none"
                }
              },
              {
                key: "versionHistory",
                text: "Version History",
                cacheKey: "myCacheKey",
                style: {
                  display:
                    selectedItems.length === 1 &&
                    selectedItems[0].selectedItemExt !== "aspx"
                      ? "inline-block"
                      : "none"
                },
                onClick: () => setVersionHistoryForm(true)
              },
              {
                key: "brandDoc",
                text: "Brand Documents",
                cacheKey: "myCacheKey",
                style: {
                  display: checkMultiForAspx(selectedItems)
                },
                onClick: () => setDocumentBrandForm(true)
              }
            ]
          }}
          disabled={!selectedItems || selectedItems.length === 0}
          iconProps={{ iconName: "MoreVertical" }}
          styles={{
            root: {
              marginLeft: 10
            },
            icon: { color: "#808080", fontSize: 19 },
            iconHovered: { color: "#808080" },
            menuIcon: { display: "none" }
          }}
        />
        <div className="calloutArea" ref={contextualMneuDialogRef}>
          {isCopyLinkDialog && (
            <Callout
              gapSpace={0}
              target={contextualMneuDialogRef.current}
              onDismiss={() => setIsCopyLinkDialog(false)}
              setInitialFocus={true}
              isBeakVisible={false}
              directionalHint={DirectionalHint.bottomCenter}
            >
              <iframe
                style={{ width: "350px", height: "250px" }}
                src={copyLink(
                  selectedListId,
                  selectedItems[0].selectedItemId.toString()
                )}
                frameBorder={0}
              />
            </Callout>
          )}

          {isShareLinkDialog && (
            <Callout
              gapSpace={0}
              target={contextualMneuDialogRef.current}
              onDismiss={() => setIsShareLinkDialog(false)}
              setInitialFocus={true}
              isBeakVisible={false}
              directionalHint={DirectionalHint.bottomCenter}
            >
              <ShareLinkForm
                listId={selectedListId}
                itemId={selectedItems[0].selectedItemId.toString()}
              />
            </Callout>
          )}
        </div>

        {isAlerMeDialog && (
          <AlertMeForm
            isDialog={isAlerMeDialog}
            onDismiss={() => setAlerMeDialog(false)}
            link={alertMeLink(selectedListId, selectedItemId.toString())}
          />
        )}

        {isFeedbackForm && (
          <FeedbackForm
            isOpen={isFeedbackForm}
            onCloseForm={() => setFeedbackForm(false)}
            feedbackFormSettings={feedbackForm}
            docId={docId}
            stream={stream}
            selectedItems={selectedItems}
          />
        )}

        {isVersionHistoryForm && (
          <VersionHistoryForm
            onDismiss={() => setVersionHistoryForm(false)}
            isDialog={isVersionHistoryForm}
            link={versionHistoryLink(
              selectedListId,
              selectedItems[0].selectedItemId.toString()
            )}
          />
        )}

        {isDocumentBrandForm && (
          <DocumentBrandForm
            isOpen={isDocumentBrandForm}
            onCloseForm={() => _handleOnClose()}
            selectedItems={selectedItems}
            selectedList={selectedListId}
            setSelectedItems={setSelectedItems}
            setClearSelection={setClearSelection}
          />
        )}
      </div>
    );
  }
);
