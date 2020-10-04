import * as React from "react";
import {
  ICommandBarItemProps,
  IContextualMenuItemProps
} from "office-ui-fabric-react";
import { exportToExcel } from "../utils/exportToExcel";
import { getDocTypeIcon } from "../styles/Icons";
import { getOpentInLink } from "../utils/openInLink";
import { dowloadSingleFile } from "../utils/dowloadSingleFile";
import { getZippedFiles } from "../utils/getZippedFiles";
import { ISelectedItem } from "../interfaces/ISelectedItem";
import { manageAlertLink } from "../utils/manageAlertLink";

const getDisplayAttr = (selectedItems: ISelectedItem[]) => {
  return selectedItems.length === 1 &&
    //selectedItems[0].selectedItemExt !== "aspx" &&
    selectedItems[0].selectedItemExt !== "url"
    ? "inline-block"
    : "none";
};

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

export const menuItems = (
  listId: string,
  viewId: string,
  onForm: (value: boolean) => void,
  activateFeedbackForm: boolean,
  selectedItems: ISelectedItem[]
): ICommandBarItemProps[] => [
  {
    key: "exportTo",
    text: "Export to Excel",
    cacheKey: "myCacheKey",
    style: {
      backgroundColor: "#fff"
    },
    href: exportToExcel(listId, viewId),
    iconProps: { iconName: "ExcelLogo" }
  },
  {
    key: "manageAlerts",
    text: "Manage My Alerts",
    iconProps: { iconName: "PeopleAlert" },
    href: manageAlertLink(),
    target: "_blank",
    style: {
      backgroundColor: "#fff"
    },
    ["data-interception"]: "off"
  },
  {
    key: "feedback",
    text: "Feedback",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "Feedback" },
    style: {
      backgroundColor: "#fff",
      display:
        activateFeedbackForm && selectedItems.length < 2
          ? "inline-block"
          : "none"
    },
    onClick: () => onForm(true)
  }
];

export const activeMenuItems = (
  selectedListInternalName: string,
  activateFeedbackForm: boolean,
  onAlertMe: (value: boolean) => void,
  onCopyLink: (value: boolean) => void,
  onShareLink: (value: boolean) => void,
  onDocumentBrandForm: (value: boolean) => void,
  onForm: (value: boolean) => void,
  selectedItems: ISelectedItem[],
  onVersionHistory: (value: boolean) => void
): ICommandBarItemProps[] => [
  {
    key: "open",
    text: "Open",
    cacheKey: "myCacheKey",
    onRenderIcon: () => getDocTypeIcon(selectedItems[0].selectedItemExt),
    style: {
      backgroundColor: "#fff",
      display: getDisplayAttr(selectedItems)
    },
    onClick: () => onShareLink(true),
    subMenuProps: {
      contextualMenuItemAs: (props: IContextualMenuItemProps) => {
        return (
          <>
            {getDocTypeIcon(selectedItems[0].selectedItemExt)}
            <span>{props.item.text}</span>
          </>
        );
      },
      styles: { container: { overflow: "hidden" } },
      items: [
        {
          key: "openInbrowser",
          text: "Open in browser",
          href:
            selectedItems[0].selectedItemExt !== "aspx"
              ? selectedItems[0].selectedItemUrlOpenInBrowser
              : getOpentInLink(
                  selectedItems[0].selectedItemExt,
                  selectedListInternalName,
                  selectedItems[0].selectedItemName
                ),
          target: "_blank",
          ["data-interception"]: "off"
        },
        {
          key: "openInApp",
          text: "Open in app",
          href: getOpentInLink(
            selectedItems[0].selectedItemExt,
            selectedListInternalName,
            selectedItems[0].selectedItemName
          ),
          style:
            selectedItems[0].selectedItemExt === "aspx"
              ? { display: "none" }
              : { display: "inline-block" }
        }
      ]
    }
  },

  // {
  //   key: "share",
  //   text: "Share",
  //   cacheKey: "myCacheKey",
  //   iconProps: { iconName: "Share" },
  //   style: {
  //     backgroundColor: "#fff",
  //     display:
  //       selectedItems.length === 1 &&
  //       selectedItems[0].selectedItemExt !== "aspx"
  //         ? "inline-block"
  //         : "none"
  //   },
  //   onClick: () => onShareLink(true)
  // },

  {
    key: "copyLink",
    text: "Copy link",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "Link" },
    style: {
      backgroundColor: "#fff",
      display: selectedItems.length === 1 ? "inline-block" : "none"
    },
    onClick: () => onCopyLink(true)
  },
  {
    key: "download",
    text: "Download",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "Download" },
    style: {
      backgroundColor: "#fff",
      display:
        selectedItems.length === 1 &&
        selectedItems[0].selectedItemExt === "aspx"
          ? "none"
          : "inline-block"
    },
    href: selectedItems.length === 1 && dowloadSingleFile(selectedItems[0]),
    onClick:
      selectedItems.length > 1
        ? async () => await getZippedFiles(selectedItems)
        : () => null
  },

  {
    key: "alert",
    text: "Alert Me",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "AlertSolid" },
    style: {
      backgroundColor: "#fff",
      display:
        selectedItems.length === 1 &&
        selectedItems[0].selectedItemExt !== "aspx"
          ? "inline-block"
          : "none"
    },
    onClick: () => onAlertMe(true)
  },
  {
    key: "manageAlerts",
    text: "Manage My Alerts",
    iconProps: { iconName: "PeopleAlert" },
    href: manageAlertLink(),
    target: "_blank",
    style: {
      backgroundColor: "#fff"
    },
    ["data-interception"]: "off"
  },

  {
    key: "feedback",
    text: "Feedback",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "Feedback" },
    style: {
      backgroundColor: "#fff",
      display:
        activateFeedbackForm && selectedItems.length === 1
          ? "inline-block"
          : "none"
    },
    onClick: () => onForm(true)
  },

  {
    key: "versionHistory",
    text: "Version History",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "History" },
    style: {
      backgroundColor: "#fff",
      display:
        selectedItems.length === 1 &&
        selectedItems[0].selectedItemExt !== "aspx"
          ? "inline-block"
          : "none"
    },
    onClick: () => onVersionHistory(true)
  },
  {
    key: "brandDoc",
    text: "Brand Documents",
    cacheKey: "myCacheKey",
    iconProps: { iconName: "Tag" },
    style: {
      backgroundColor: "#fff",
      display: checkMultiForAspx(selectedItems)
    },
    onClick: () => onDocumentBrandForm(true)
  }
];
