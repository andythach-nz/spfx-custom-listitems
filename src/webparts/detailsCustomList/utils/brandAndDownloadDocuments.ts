import SharePointService from "../services/SharePointService";
import * as FileSaver from "file-saver";
import {
  HttpClient,
  IHttpClientOptions,
  HttpClientResponse
} from "@microsoft/sp-http";
import { Dialog } from "@microsoft/sp-dialog";
import { getSavedDocuments, getBrandingFormProperty } from "./getLookUpFields";
import * as React from "react";
import { IBrandingFormProps } from "../interfaces/IBrandingFormProps";

export const brandAndDownloadDocuments = async (
  Brand: string,
  items: string[], //doc ids
  ListId: string, //process area for now, make it array
  success: string,
  failed: string,
  brandingValues: IBrandingFormProps
): Promise<void> => {
  console.log("brandingValues", brandingValues);
  const flowURL =
    "https://prod-19.australiasoutheast.logic.azure.com:443/workflows/fe4da41d9a1b4c0f86d040ba589e1c9e/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=kKBUjx6viUY-GRRM5eIMyMrYgSfYeMaozVZ1i0uIeSw";

  const requestHeaders: Headers = new Headers();
  requestHeaders.append("Content-type", "application/json");

  let data = {
    Brand: Brand,
    ListItems: items,
    ListId: ListId
  };

  let requestoptions: IHttpClientOptions = {
    method: "POST",
    headers: requestHeaders,
    body: JSON.stringify(data)
  };

  await SharePointService.context.httpClient
    .post(flowURL, HttpClient.configurations.v1, requestoptions)
    .then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          let zipBase64 = result.FileBase64;
          fetch("data:application/zip;base64," + zipBase64)
            .then(res => res.blob())
            .then(blob => {
              FileSaver.saveAs(blob, "Documents.zip");
            });
        });
        try {
          SharePointService.pnp_addItem(
            "SubmittedBrandingDocument",
            brandingValues
          );
          getSavedDocuments().then(obj => {
            obj.map(i => {
              if (
                i.Email === SharePointService.context.pageContext.user.email
              ) {
                SharePointService.pnp_deleteItem("SavedBrandingDocument", i.ID);
              }
            });
          });
        } catch (error) {
          console.log(error);
        }
      } else {
        Dialog.alert(failed);
      }
    });
};

// This function iterate and call the brandAndDownloadDocuments() depending the list of libraries selected.
export const iteration = async (
  brand: string,
  documentsObjNotation: string,
  success: string,
  failed: string,
  brandingValues: IBrandingFormProps
): Promise<void> => {
  console.log("documentsObjNotation", documentsObjNotation);

  let Obj = JSON.parse(documentsObjNotation);
  console.log("Obj", Obj);

  let tempArray = [],
    uniqueListIds = [];

  for (let i = 0; i < Obj.length; i++) {
    if (tempArray[Obj[i].selectedList]) continue;
    tempArray[Obj[i].selectedList] = true;
    uniqueListIds.push(Obj[i].selectedList);
  }
  console.log("uniqueListIds", uniqueListIds);

  setTimeout(() => {
    uniqueListIds.map((listID, index) => {
      let docIds = [];
      for (let i = 0; i < Obj.length; i++) {
        if (
          //only word docments can be branded
          (listID === Obj[i].selectedList &&
            Obj[i].selectedItemName.includes(".doc")) ||
          Obj[i].selectedItemName.includes(".dot")
        ) {
          docIds.push(Obj[i].selectedItemId);
        }
      }
      console.log("docIds", docIds);
      brandAndDownloadDocuments(
        brand,
        docIds,
        listID,
        success,
        failed,
        brandingValues
      );
    });
  }, 5000);
};
