import { LocalStrings } from "./consts/LocalStrings";
import { IInputs } from "./generated/ManifestTypes";
import { EntityMetadata } from "./Interfaces";
import { ActivityType } from "./Interfaces";

export async function getEntityMetadata(
  context: ComponentFramework.Context<IInputs>
): Promise<EntityMetadata | null> {
  if (!context || !(context as any).page) {
    console.warn("Component Framework context is not available. (utils)");
    return null;
  }

  const dynamicsUrl = (context as any).page.getClientUrl();
  if (!dynamicsUrl) {
    console.error("Unable to retrieve client URL.");
    return null;
  }

  const entityName = (context as any).page.entityTypeName;
  if (!entityName) {
    console.error("Unable to retrieve entity type name.");
    return null;
  }
  const apiUrl = `${dynamicsUrl}/api/data/v9.0/EntityDefinitions(LogicalName='${entityName}')?$select=SchemaName,LogicalCollectionName`;

  const entityId = (context as any).page.entityId;
  if (!entityId) {
    console.error("Unable to retrieve entity ID.");
    return null;
  }

  try {
    const response = await fetch(apiUrl, {
      headers: {
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        Accept: "application/json",
        "Content-Type": "application/json; charset=utf-8",
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return {
      schemaName: data.SchemaName.toLowerCase(),
      logicalCollectionName: data.LogicalCollectionName,
      clientUrl: dynamicsUrl,
      entityId: entityId,
    };
  } catch (error) {
    console.error("Error fetching entity metadata:", error);
    return null;
  }
}

export function isPDF(mime: string) {
  return mime === "application/pdf";
}

export function isExcel(mime: string) {
  return (
    mime ===
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
    mime === "application/vnd.ms-excel" ||
    mime === "text/csv"
  );
}

export function isImage(mime: string) {
  const base64Prefix = "data:image/";
  if (mime.startsWith(base64Prefix)) {
    const parts = mime.split(base64Prefix);
    if (parts.length > 2) {
      mime = base64Prefix + parts[1];
    }
  }
  return mime.startsWith("image/");
}

export function createDataUri(mimetype: string, base64: string): string {
  return `data:${mimetype};base64,${base64}`;
}

export function isActivityType(schemaName: string): schemaName is ActivityType {
  return [
    "email",
    "phonecall",
    "appointment",
    "task",
    "fax",
    "letter",
    "serviceappointment",
    "campaignresponse",
    "campaignactivity",
    "bulkoperation",
    "socialactivity",
    "recurringappointmentmaster",
    "appointmentrecurrence",
  ].includes(schemaName);
}
export function getControlValue(
  context: ComponentFramework.Context<any>,
  parameter: string
) {
  return context.parameters[parameter]?.raw || "";
}
const delay = (ms: number): Promise<void> => {
  return new Promise((resolve) => setTimeout(resolve, ms));
};

export async function focusSPDocumentsAndRestore() {
  const originalTab = getFocusedTab();
  const navItem = Xrm.Page.ui.navigation.items.get("navSPDocuments");
  
  if (navItem) {
    navItem.setFocus();
    await delay(1000);
    setTimeout(function () {
      if (originalTab) {
        originalTab.setFocus();
        window.location.reload();
      }
    }, 1000);
  } else {
    console.error("SP Documents navigation item not found.");
  }
}

function getFocusedTab() {
  const tabs = Xrm.Page.ui.tabs;

  if (tabs.getLength() > 0) {
    const firstTab = tabs.get(0);
    return firstTab;
  } else {
    return null;
  }
}

  type Values<T> = T extends string ? T : { [K in keyof T]: Values<T[K]> }[keyof T];
  type LocalizationKey = Values<typeof LocalStrings>;

  export const getLocalString = (
  context: ComponentFramework.Context<any>,
  key: LocalizationKey
): string => context.resources.getString(key);

export function b64toBlob(b64Data: string, contentType = '', sliceSize = 512) {
  if (b64Data.startsWith("data:application/pdf;base64,")) {
    b64Data = b64Data.replace("data:application/pdf;base64,", "");
  }

  const byteCharacters = atob(b64Data);
  const byteArrays = [];

  for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
    const slice = byteCharacters.slice(offset, offset + sliceSize);
    const byteNumbers = new Array(slice.length);
    for (let i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i);
    }
    byteArrays.push(new Uint8Array(byteNumbers));
  }

  return new Blob(byteArrays, { type: contentType });
}

