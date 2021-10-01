import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";
import { reject } from "lodash";

export class SPOperations {
  public GetAllLists(context: WebPartContext): Promise<IDropdownOption[]> {
    let _spLists: IDropdownOption[] = [];
    let apiURL =
      context.pageContext.web.absoluteUrl + "/_api/web/lists?select=Title";
    return new Promise<IDropdownOption[]>((resolve, reject) => {
      context.spHttpClient
        .get(apiURL, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            results.value.map((result: any) => {
              _spLists.push({ key: result.Title, text: result.Title });
            });
          });
          resolve(_spLists);
        });
    });
  }

  public CreateListItem(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let apiURL =
      context.pageContext.web.absoluteUrl +
      `/_api/web/lists/getByTitle('${listTitle}')/items`;
    const body = JSON.stringify({ Title: "Second Item Created" });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "Application/json;odata=nometadata",
        "Content-Type": "Application/json;odata=nometadata",
        "odata-version": "",
      },
      body: body,
    };
    return new Promise<string>((resolve, reject) => {
      context.spHttpClient
        .post(apiURL, SPHttpClient.configurations.v1, options)
        .then((response: SPHttpClientResponse) => {
          response.json().then((result: any) => {
            resolve(`Item with ID ${result.Id} created succesfully`);
          });
        });
    });
  }

  public UpdateListItem(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let apiURL =
      context.pageContext.web.absoluteUrl +
      `/_api/web/lists/getByTitle('${listTitle}')/getItemById(3)`;
    const body = JSON.stringify({ Title: "Updated Title" });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "Application/json;odata=nometadata",
        "Content-Type": "Application/json;odata=nometadata",
        "odata-version": "",
        "IF-Match": "*",
        "X-HTTP-METHOD": "MERGE",
      },
      body: body
    };
    return new Promise<string>((resolve, reject) => {
      context.spHttpClient
        .post(apiURL, SPHttpClient.configurations.v1, options)
        .then(
          () => {
            resolve("successfully updated");
          },
          (err: any) => reject("Error Occured")
        );
    });
  }
  public DeleteListItem(
    context: WebPartContext,
    listTitle: string
  ): Promise<string> {
    let apiURL =
      context.pageContext.web.absoluteUrl +
      `/_api/web/lists/getByTitle('${listTitle}')/items(2)`;

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "Application/json;odata=nometadata",
        "Content-Type": "Application/json;odata=nometadata",
        "odata-version": "",
        "IF-Match": "*",
        "X-HTTP-METHOD": "DELETE",
      },
    };
    return new Promise<string>((resolve, reject) => {
      context.spHttpClient
        .post(apiURL, SPHttpClient.configurations.v1, options)
        .then(() => {
          resolve("Successfully deleted");
        });
    });
  }
}
