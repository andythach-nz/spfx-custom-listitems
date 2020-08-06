import * as React from "react";
import {
  Modal,
  IconButton,
  getTheme,
  IIconProps,
  mergeStyleSets,
  FontWeights
} from "office-ui-fabric-react";

export interface IVersionHistoryFormProps {
  link: string;
  isDialog: boolean;
  onDismiss: () => void;
}
const theme = getTheme();
const iconButtonStyles = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: "auto",
    marginTop: "4px",
    marginRight: "2px",
    textAlign: "right"
  },
  rootHovered: {
    color: theme.palette.neutralDark
  }
};
const cancelIcon: IIconProps = { iconName: "Cancel" };
const contentStyles = mergeStyleSets({
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: "1 1 auto",
      color: theme.palette.neutralPrimary,
      display: "flex",
      alignItems: "center",
      fontWeight: FontWeights.semibold
    }
  ]
});

export const VersionHistoryForm: React.FC<IVersionHistoryFormProps> = ({
  isDialog,
  onDismiss,
  link
}): JSX.Element => {
  const [iframeVH, setIframeVH] = React.useState();
  const [aspnetVHForm, setAspnetVHForm] = React.useState();
  console.log("link", link);

  React.useEffect(() => {
    setIframeVH(iframeVH);
  }, [iframeVH]);

  return (
    <Modal
      isOpen={isDialog}
      titleAriaId="VersionHistoryForm"
      onDismiss={onDismiss}
      styles={{ scrollableContent: { maxHeight: "90vh", padding: "15px" } }}
    >
      {/*!aspnetVHForm && <Spinner size={SpinnerSize.large} />*/}
      <div className={contentStyles.header}>
        <IconButton
          styles={iconButtonStyles}
          iconProps={cancelIcon}
          ariaLabel="Close popup modal"
          onClick={onDismiss}
        />
      </div>

      <iframe
        ref={setIframeVH}
        style={{ width: "600px", height: "1150px", overflow: "scroll" }}
        src={link}
        frameBorder={0}
      />
    </Modal>
  );
};
