import * as React from "react";
import {
  Panel,
  PanelType,
  Stack,
  Image,
  Label,
  Separator,
  IStackTokens,
  Dropdown,
  IDropdownOption,
  Toggle,
  PrimaryButton,
  Link
} from "office-ui-fabric-react";
import SharePointService from "../services/SharePointService";
import { IBrandingFormProps } from "../interfaces/IBrandingFormProps";
const formImage: string = require("../images/TheDownerStandard.jpg");
import { iteration } from "../utils/brandAndDownloadDocuments";
import { Dialog } from "@microsoft/sp-dialog";
import {
  getBrandingOptions,
  getSavedDocuments
} from "../utils/getLookUpFields";
import { formLabels } from "../utils/getBrandingFormLabels";
import { getUniqueValues } from "../utils/getUniqueValues";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { ISelectedItem } from "../interfaces/ISelectedItem";
import { ILabelProps } from "../interfaces/ILabelProps";
import { Timeout } from "../utils/Timeout";
const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 2
};
export const BrandingValues: IBrandingFormProps = {
  Title: "",
  Email: "",
  Department: "",
  Brand: "",
  NominatedContentAdmin: "",
  ReadAndAgreed: "",
  SavedDocumentObject: ""
};

interface IDocumentBrandFormProps {
  isOpen: boolean;
  onCloseForm: () => void;
  selectedItems: ISelectedItem[];
  selectedList: string;
  setSelectedItems: (selectedItems: ISelectedItem[]) => void;
  setClearSelection: (value: boolean) => void;
}

const _onRenderHeader = (): JSX.Element => {
  return (
    <Stack verticalAlign="center">
      <Image
        src={formImage}
        width={360}
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

export const DocumentBrandForm: React.FC<IDocumentBrandFormProps> = ({
  isOpen,
  onCloseForm,
  selectedItems,
  selectedList,
  setSelectedItems,
  setClearSelection
}): JSX.Element => {
  const [userDepartment, setUserDepartment] = React.useState("");
  const [brandingValues, setBrandingValues] = React.useState<
    IBrandingFormProps
  >(BrandingValues);
  const [brandOptions, setBrandOptions] = React.useState<IDropdownOption[]>();
  const [dataFetchComplete, setDataFetchComplete] = React.useState<boolean>(
    false
  );
  const [formLabel, setFormLabel] = React.useState<ILabelProps>(formLabels);

  React.useEffect(() => {
    setFormLabel(formLabels);
    setDataFetchComplete(true);
  }, []);

  React.useEffect(() => {
    setFormLabel(formLabels);
    getBrandingOptions().then(r => {
      let tempArray = [];
      r.map(item => {
        tempArray.push(item.Brand);
      });
      const uniqueValues = tempArray.filter(getUniqueValues);
      const brandOption = uniqueValues.map(brand => ({
        key: brand,
        text: brand
      }));
      setBrandOptions(brandOption);
    });
  }, []);

  React.useEffect(() => {
    console.log("selectedItems", selectedItems);
    let tempSelectedItems = [];
    if (selectedItems.length > 0) {
      selectedItems.map(selectedItem => {
        tempSelectedItems.push({ ...selectedItem, selectedList: selectedList });
      });
    }

    getSavedDocuments().then(savedItems => {
      if (savedItems.length > 0) {
        savedItems.map((savedItem, index, arr) => {
          if (
            savedItem.Email === SharePointService.context.pageContext.user.email
          ) {
            const objectFromString = JSON.parse(savedItem.SavedDocumentObject);
            if (index + 1 === arr.length) {
              objectFromString.map(s => {
                console.log("s", s);
                tempSelectedItems.push(s);
              });
            }
          }
        });
      }
    });
    setSelectedItems(tempSelectedItems);
    console.log("tempSelectedItems", tempSelectedItems);
    SharePointService.pnp_getUserProfileProperty(
      "i:0#.f|membership|" +
        SharePointService.context.pageContext.user.loginName,
      "Department"
    ).then(dept => {
      setUserDepartment(dept);
      setBrandingValues({
        Title: SharePointService.context.pageContext.user.displayName,
        Email: SharePointService.context.pageContext.user.email,
        Department: dept,
        Brand: "",
        NominatedContentAdmin: "",
        ReadAndAgreed: "No",
        SavedDocumentObject: JSON.stringify(tempSelectedItems)
      });
    });
  }, []);

  const _getPeoplePickerItems = (items: any[]) => {
    setBrandingValues({
      ...brandingValues,
      NominatedContentAdmin: items[0].secondaryText
    });
  };

  const _handleOnChange = (
    e: React.FormEvent<HTMLInputElement>,
    inputValue: IDropdownOption,
    index: any
  ) => {
    const currentId = e.target["id"] as string;
    if (currentId === "brandingSelectionKeys") {
      setBrandingValues({ ...brandingValues, Brand: inputValue.text });
    }
  };

  const _generalUserInformation = (): JSX.Element => {
    return (
      <div>
        <div style={{ display: "flex" }}>
          <Label style={{ fontWeight: 500, paddingRight: "5px" }}>
            {formLabel.Name}
            {":"}
          </Label>
          <Label>
            {SharePointService.context.pageContext.user.displayName}
          </Label>
        </div>
        <div style={{ display: "flex" }}>
          <Label style={{ fontWeight: 500, paddingRight: "5px" }}>
            {formLabel.Email}
            {":"}
          </Label>
          <Label>{SharePointService.context.pageContext.user.email}</Label>
        </div>
        <div style={{ display: "flex" }}>
          <Label style={{ fontWeight: 500, paddingRight: "5px" }}>
            {formLabel.Department}
            {":"}
          </Label>
          <Label>{userDepartment}</Label>
        </div>
      </div>
    );
  };

  const _onDisclaimer = (): JSX.Element => {
    return (
      <div
        style={{
          backgroundColor: "rgb(197, 227, 244)",
          padding: "10px",
          marginTop: "10px"
        }}
      >
        <p style={{ fontWeight: "bold", fontSize: "14px" }}>
          Please be aware of the following:
        </p>
        <p>{formLabel.DisclaimerParagraph}</p>
        <ul>
          <div
            dangerouslySetInnerHTML={{ __html: formLabel.DisclaimerBullets }}
          ></div>
          <div style={{ textAlign: "center", marginTop: "5px" }}>
            <label style={{ fontWeight: "bold" }}>I have read and agree:</label>
            <Toggle
              id="disclaimerToggle"
              onText="Yes"
              offText="No"
              onChange={(e: any, checked?: boolean) =>
                checked
                  ? setBrandingValues({
                      ...brandingValues,
                      ReadAndAgreed: "Yes"
                    })
                  : setBrandingValues({
                      ...brandingValues,
                      ReadAndAgreed: "No"
                    })
              }
            />
          </div>
        </ul>
      </div>
    );
  };

  const _saveForm = async (e: any): Promise<void> => {
    e.preventDefault();
    try {
      getSavedDocuments().then(obj => {
        if (obj.length > 0) {
          for (let i = 0; i < obj.length; i++) {
            if (
              obj[i].Email === SharePointService.context.pageContext.user.email
            ) {
              SharePointService.pnp_updateItem(
                "SavedBrandingDocument",
                obj[i].ID,
                brandingValues
              );
              break;
            } else {
              SharePointService.pnp_addItem(
                "SavedBrandingDocument",
                brandingValues
              );
            }
          }
        } else {
          SharePointService.pnp_addItem(
            "SavedBrandingDocument",
            brandingValues
          );
        }
      });
      Dialog.alert(formLabel.SavedMessage);
      onCloseForm();
    } catch (error) {
      onCloseForm();
      throw error;
    }
  };

  const _submitForm = async (e: any): Promise<void> => {
    e.preventDefault();
    try {
      Dialog.alert(formLabel.SubmitSuccess);
      iteration(
        brandingValues.Brand,
        brandingValues.SavedDocumentObject,
        formLabel.SubmitSuccess,
        formLabel.SubmitFailed,
        brandingValues
      );
      onCloseForm();
    } catch (error) {
      onCloseForm();
      throw error;
    }
  };

  const _removeSelectedDocuments = (e: any) => {
    e.preventDefault();
    const filtered = selectedItems.filter((el, index) => {
      return index != e.target.value;
    });
    setSelectedItems(filtered);
    let savedObj = {
      Title: SharePointService.context.pageContext.user.displayName,
      Email: SharePointService.context.pageContext.user.email,
      Department: userDepartment,
      Brand: brandingValues.Brand,
      NominatedContentAdmin: brandingValues.NominatedContentAdmin,
      ReadAndAgreed: brandingValues.ReadAndAgreed,
      SavedDocumentObject: JSON.stringify(filtered)
    };
    setBrandingValues(savedObj);
    getSavedDocuments().then(obj => {
      if (obj.length > 0) {
        for (let i = 0; i < obj.length; i++) {
          if (
            obj[i].Email === SharePointService.context.pageContext.user.email
          ) {
            try {
              SharePointService.pnp_updateItem(
                "SavedBrandingDocument",
                obj[i].ID,
                savedObj
              );
            } catch (err) {
              throw err;
            }
          }
        }
      }
    });
  };

  const _removeAllSelectedDocuments = () => {
    setSelectedItems([]);
    setClearSelection(true);
    const clearedObj = {
      Title: SharePointService.context.pageContext.user.displayName,
      Email: SharePointService.context.pageContext.user.email,
      Department: userDepartment,
      Brand: brandingValues.Brand,
      NominatedContentAdmin: brandingValues.NominatedContentAdmin,
      ReadAndAgreed: brandingValues.ReadAndAgreed,
      SavedDocumentObject: "[]"
    };
    getSavedDocuments().then(obj => {
      if (obj.length > 0) {
        for (let i = 0; i < obj.length; i++) {
          if (
            obj[i].Email === SharePointService.context.pageContext.user.email
          ) {
            try {
              SharePointService.pnp_updateItem(
                "SavedBrandingDocument",
                obj[i].ID,
                clearedObj
              );
            } catch (err) {
              throw err;
            }
          }
        }
      }
    });
  };

  const _onRenderFooterContent = () => {
    return (
      <Stack horizontal horizontalAlign="end">
        <PrimaryButton
          onClick={_saveForm}
          text="Save form"
          disabled={!brandingValues.NominatedContentAdmin}
        />
        <PrimaryButton
          style={{ marginLeft: "5px" }}
          onClick={_submitForm}
          text="Submit"
          disabled={
            !brandingValues.NominatedContentAdmin ||
            brandingValues.ReadAndAgreed === "No" ||
            brandingValues.Brand === "My brand is not listed"
          }
        />
      </Stack>
    );
  };

  const onDismiss = () => {
    setSelectedItems([]);
    onCloseForm();
    setClearSelection(true);
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.custom}
      customWidth="520px"
      onDismiss={onDismiss}
      onRenderHeader={_onRenderHeader}
      onRenderFooterContent={_onRenderFooterContent}
    >
      <form onSubmit={_submitForm}>
        <Stack tokens={itemAlignmentsStackTokens}>
          <Label>
            <p>
              Thank you for taking the time to submit your request. Your name,
              email address and request details will be sent to the relevant
              owner for consideration.
            </p>
            <p>
              Note: For all IT related issues contact:
              <br />
              Australia - 1300 333 000
              <br />
              New Zealand - 0800 156 666
              <br /> Spotless- AU: 1300 333 000, NZ: 0800 487 768
            </p>
          </Label>
        </Stack>

        {_generalUserInformation()}

        <Dropdown
          placeholder="Select an option"
          label={formLabel.BrandingRequired}
          id="brandingSelectionKeys"
          options={brandOptions}
          onChange={_handleOnChange}
        />

        {_onDisclaimer()}

        <PeoplePicker
          context={SharePointService.context}
          titleText={formLabel.NominatedContentOwner}
          personSelectionLimit={1}
          //groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
          showtooltip={true}
          isRequired={true}
          selectedItems={_getPeoplePickerItems}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
        />

        <div style={{ marginTop: "10px" }}>
          <label style={{ fontWeight: "bold" }}>Documents requested:</label>
          <ul style={{ marginLeft: "20px" }}>
            {dataFetchComplete &&
              Timeout(500) &&
              selectedItems.map((file, index) => {
                return (
                  <li
                    style={{
                      listStyle: "square"
                    }}
                  >
                    {file.selectedItemName}{" "}
                    <Link value={index} onClick={_removeSelectedDocuments}>
                      remove
                    </Link>
                  </li>
                );
              })}
          </ul>
          <Link
            style={
              selectedItems.length < 0
                ? { display: "none" }
                : { display: "inline-block" }
            }
            onClick={_removeAllSelectedDocuments}
          >
            Remove All
          </Link>
        </div>
      </form>
    </Panel>
  );
};
