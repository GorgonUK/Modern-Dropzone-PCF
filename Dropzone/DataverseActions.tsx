import { IInputs } from "./generated/ManifestTypes";
import {GenericActionResponse, NoteView} from "./Interfaces"
import { getEntityMetadata, isActivityType } from "./utils";

export async function createRelatedNote(
  context: ComponentFramework.Context<IInputs>,
  fileName: string,
  base64: string,
  fileSize: number,
  mimeType: string
) {
  const entity = "annotation";
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return { success: false, message: "Failed to retrieve entity metadata" };
  }
  const objectId = `objectid_${metadata.schemaName}@odata.bind`;
  const record = {
    "isdocument": true,
    "filename": fileName,
    "subject": fileName,
    "filesize": fileSize,
    "mimetype": mimeType,
    "documentbody": base64.split(',')[1],
    [objectId]: `/${metadata?.logicalCollectionName}(${metadata.entityId})`
  };

  try {
    const result = await context.webAPI.createRecord(entity, record);
    return { success: true, message: "Note created successfully", noteId: result.id };
  } catch (error) {
    console.error("Error creating note:", error);
    return { success: false, message: `Error creating note: ${(error as any).message}` };
  }
}

export async function getRelatedNotes(
  context: ComponentFramework.Context<IInputs>
) {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return { success: false, message: "Failed to retrieve entity metadata", data: [] };
  }
  const filter = `_objectid_value eq ${metadata.entityId}`;

  try {
    const result = await context.webAPI.retrieveMultipleRecords('annotation', `?$filter=${filter}&$orderby=createdon desc`);
    const filteredEntities = result.entities.filter((entity: any) => entity.objecttypecode === metadata.schemaName);
    return { success: true, data: filteredEntities };
  } catch (error) {
    console.error("Error fetching notes:", error);
    return { success: false, message: `Error fetching notes: ${(error as any).message}`, data: [] };
  }
}

export async function deleteRelatedNote(
  context: ComponentFramework.Context<IInputs>,
  noteId: string
) {
  try {
    await context.webAPI.deleteRecord('annotation', noteId);
    return { success: true, message: "Note deleted successfully" };
  } catch (error) {
    console.error("Error deleting note:", error);
    return { success: false, message: `Error deleting note: ${(error as any).message}` };
  }
}

export async function updateRelatedNote(
  context: ComponentFramework.Context<IInputs>,
  noteId: string,
  filename: string,
): Promise<{ success: boolean; message: string }> {
  const data = {
    filename: filename
  };

  try {
    await context.webAPI.updateRecord('annotation', noteId, data);
    return {
      success: true,
      message: "Note updated successfully."
    };
  } catch (error) {
    console.error("Failed to update note:", error);
    return {
      success: false,
      message: `Error updating note: ${(error as Error).message}`
    };
  }
}


export async function getNoteViews(
  context: ComponentFramework.Context<IInputs>
): Promise<{ success: boolean; message?: string; data: any[] }> {
  try {
    const select = "$select=savedqueryid,name,fetchxml";
    const filter = "$filter=returnedtypecode eq 'annotation' and fetchxml ne null";

    const res = await context.webAPI.retrieveMultipleRecords(
      "savedquery",
      `?${select}&${filter}`
    );

    const views: NoteView[] = res.entities.map((e: any) => ({
      savedqueryid: e.savedqueryid,
      name:         e.name,
      fetchxml:     e.fetchxml
    }));

    return { success: true, data: views };

  } catch (err: any) {
    console.error("Error fetching note views:", err);
    return {
      success: false,
      message: `Error fetching note views: ${err.message}`,
      data: []
    };
  }
}



export async function getSharePointLocations(context: ComponentFramework.Context<IInputs>): Promise<{ name: string, sharepointdocumentlocationid:string }[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  const fetchXml = `
  <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
  <entity name="sharepointdocumentlocation">
    <attribute name="statecode" />
    <attribute name="name" />
    <attribute name="sharepointdocumentlocationid" />
    <attribute name="servicetype" />
    <attribute name="parentsiteorlocation" />
    <order attribute="modifiedon" descending="true" />
    <filter type="and">
      <condition attribute="statuscode" operator="eq" value="1" />
      <condition attribute="statecode" operator="eq" value="0" />
    </filter>
    <filter type="and">
      <filter type="and">
        <condition attribute="regardingobjectid" operator="eq" value="${metadata.entityId}" />
      </filter>
      <filter type="and">
        <condition attribute="servicetype" operator="in">
          <value>0</value>
        </condition>
      </filter>
    </filter>
  </entity>
</fetch>
  `;
  const encodedFetchXml = encodeURIComponent(fetchXml);
  const url = `${metadata.clientUrl}/api/data/v9.0/sharepointdocumentlocations?fetchXml=${encodedFetchXml}`;

  
  const headers = {
    "Accept": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Prefer": "odata.include-annotations=OData.Community.Display.V1.FormattedValue"
  };

  const response = await fetch(url, {
    method: "GET",
    headers: headers
  });

  if (!response.ok) {
    throw new Error(`Error fetching SharePoint document locations: ${response.statusText}`);
  }
let data = await response.json();
  return data.value.map((item: any) => ({
    name: item.name,
    sharepointdocumentlocationid: item.sharepointdocumentlocationid,
    parentsiteid: item._parentsiteorlocation_value
  }));
}
export async function getSharePointData(
  context: ComponentFramework.Context<IInputs>,
  selectedDocumentLocation?:string,
  folderPath?: string,
  selectedDocumentLocationName?:string
): Promise<any[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  const isActivity = isActivityType(metadata.schemaName)
  let filters = '';
  let folderFilter = '';
  if (folderPath) {
    folderFilter = `
      <filter type="and">
        <condition attribute="relativelocation" operator="eq" value="${folderPath}"/>
      </filter>`;
  }
  filters = `
      <filter type="and">
         <filter type="and">
            <condition attribute="locationid" operator="eq" value="${selectedDocumentLocation}" />
         </filter>
         <filter type="and">
            <condition attribute="locationname" operator="eq" value="${selectedDocumentLocationName}" />
         </filter>
         <filter type="and">
            <condition attribute="servicetype" operator="eq" value="0" />
         </filter>
         ${folderFilter}
      </filter>`
  
  const fetchXml = `
    <fetch distinct="false" mapping="logical" returntotalrecordcount="true" no-lock="false">
      <entity name="sharepointdocument">
        <attribute name="documentid"/>
        <attribute name="fullname"/>
        <attribute name="relativelocation"/>
        <attribute name="sharepointcreatedon"/>
        <attribute name="ischeckedout"/>
        <attribute name="filetype"/>
        <attribute name="modified"/>
        <attribute name="sharepointmodifiedby"/>
        <attribute name="servicetype"/>
        <attribute name="absoluteurl"/>
        <attribute name="title"/>
        <attribute name="author"/>
        <attribute name="sharepointdocumentid"/>
        <attribute name="readurl"/>
        <attribute name="editurl"/>
        <attribute name="locationid"/>
        <attribute name="iconclassname"/>
        <attribute name="locationname"/>
        <attribute name="filesize"/>
        ${filters}
        <order attribute="relativelocation" descending="false"/>
        <filter>
          <condition attribute="isrecursivefetch" operator="eq" value="0"/>
        </filter>
        <link-entity name="${metadata.schemaName}" from="${isActivity?"activity":metadata.schemaName}id" to="regardingobjectid" alias="bb">
          <filter type="and">
            <condition attribute="${isActivity?"activity":metadata.schemaName}id" operator="eq" uitype="${isActivity?"activity":metadata.schemaName}" value="${metadata.entityId}"/>
          </filter>
        </link-entity>
      </entity>
    </fetch>
  `;

  const encodedFetchXml = encodeURIComponent(fetchXml);
  const url = `${metadata.clientUrl}/api/data/v9.0/sharepointdocuments?fetchXml=${encodedFetchXml}`;

  const headers = {
    "Accept": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Prefer": "odata.include-annotations=OData.Community.Display.V1.FormattedValue"
  };

  const response = await fetch(url, {
    method: "GET",
    headers: headers
  });

  if (!response.ok) {
    throw new Error(`Error fetching SharePoint data: ${response.statusText}`);
  }

  const data = await response.json();

  return data.value.map((item: any) => ({
    readurl: item.readurl,
    modified: item.modified,
    modifiedFormatted: item["modified@OData.Community.Display.V1.FormattedValue"],
    sharepointmodifiedby: item.sharepointmodifiedby,
    editurl: item.editurl,
    sharepointdocumentid: item.sharepointdocumentid,
    documentid: item.documentid,
    documentidFormatted: item["documentid@OData.Community.Display.V1.FormattedValue"],
    ischeckedout: item.ischeckedout,
    ischeckedoutFormatted: item["ischeckedout@OData.Community.Display.V1.FormattedValue"],
    sharepointcreatedon: item.sharepointcreatedon,
    sharepointcreatedonFormatted: item["sharepointcreatedon@OData.Community.Display.V1.FormattedValue"],
    locationname: item.locationname,
    iconclassname: item.iconclassname,
    absoluteurl: item.absoluteurl,
    fullname: item.fullname,
    locationid: item.locationid,
    title: item.title,
    filetype: item.filetype,
    relativelocation: item.relativelocation,
    servicetypeFormatted: item["servicetype@OData.Community.Display.V1.FormattedValue"],
    servicetype: item.servicetype,
    author: item.author,
    filesize: item.filesize
  }));
}

export async function getSharePointFolderData(context: ComponentFramework.Context<IInputs>,folderPath?: string,selectedDocumentLocation?:string|null,selectedDocumentLocationName?:string,defaultSite?:boolean): Promise<any[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  const isActivity = isActivityType(metadata.schemaName);
  let folderFilter = '';
  let filters = '';
  if (folderPath) {
    folderFilter = `
      <filter type="and">
        <condition attribute="relativelocation" operator="eq" value="${folderPath}"/>
      </filter>`;
  }
  if(selectedDocumentLocation){
    if(defaultSite){
  filters = `
      <filter type="and">
         <filter type="and">
            <condition attribute="locationid" operator="eq" value="${selectedDocumentLocation}" />
         </filter>
         <filter type="and">
            <condition attribute="locationname" operator="eq" value="${selectedDocumentLocationName}" />
         </filter>
         <filter type="and">
            <condition attribute="servicetype" operator="eq" value="0" />
         </filter>
         ${folderFilter}
      </filter>`
    }
    else{
     filters = `${folderFilter}`
    }
  }
  const fetchXml = `
    <fetch distinct="false" mapping="logical" returntotalrecordcount="true" no-lock="false">
      <entity name="sharepointdocument">
        <attribute name="documentid"/>
        <attribute name="fullname"/>
        <attribute name="relativelocation"/>
        <attribute name="sharepointcreatedon"/>
        <attribute name="ischeckedout"/>
        <attribute name="filetype"/>
        <attribute name="modified"/>
        <attribute name="sharepointmodifiedby"/>
        <attribute name="servicetype"/>
        <attribute name="absoluteurl"/>
        <attribute name="title"/>
        <attribute name="author"/>
        <attribute name="sharepointdocumentid"/>
        <attribute name="readurl"/>
        <attribute name="editurl"/>
        <attribute name="locationid"/>
        <attribute name="iconclassname"/>
        <attribute name="locationname"/>
        <attribute name="filesize"/>
        <filter>
          <condition attribute="isrecursivefetch" operator="eq" value="0"/>
        </filter>
${filters}
        <order attribute="relativelocation" descending="false"/>
        <link-entity name="${metadata.schemaName}" from="${isActivity?"activity":metadata.schemaName}id" to="regardingobjectid" alias="bb">
          <filter type="and">
            <condition attribute="${isActivity?"activity":metadata.schemaName}id" operator="eq" uitype="${isActivity?"activity":metadata.schemaName}" value="${metadata.entityId}"/>
          </filter>
        </link-entity>
      </entity>
    </fetch>
  `;

  const encodedFetchXml = encodeURIComponent(fetchXml);
  const url = `${metadata.clientUrl}/api/data/v9.0/sharepointdocuments?fetchXml=${encodedFetchXml}`;

  const headers = {
    "Accept": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Prefer": "odata.include-annotations=OData.Community.Display.V1.FormattedValue"
  };

  const response = await fetch(url, {
    method: "GET",
    headers: headers
  });

  if (!response.ok) {
    throw new Error(`Error fetching SharePoint data: ${response.statusText}`);
  }

  const data = await response.json();

  return data.value.map((item: any) => ({
    readurl: item.readurl,
    modified: item.modified,
    modifiedFormatted: item["modified@OData.Community.Display.V1.FormattedValue"],
    sharepointmodifiedby: item.sharepointmodifiedby,
    editurl: item.editurl,
    sharepointdocumentid: item.sharepointdocumentid,
    documentid: item.documentid,
    documentidFormatted: item["documentid@OData.Community.Display.V1.FormattedValue"],
    ischeckedout: item.ischeckedout,
    ischeckedoutFormatted: item["ischeckedout@OData.Community.Display.V1.FormattedValue"],
    sharepointcreatedon: item.sharepointcreatedon,
    sharepointcreatedonFormatted: item["sharepointcreatedon@OData.Community.Display.V1.FormattedValue"],
    locationname: item.locationname,
    iconclassname: item.iconclassname,
    absoluteurl: item.absoluteurl,
    fullname: item.fullname,
    locationid: item.locationid,
    title: item.title,
    filetype: item.filetype,
    relativelocation: item.relativelocation,
    servicetypeFormatted: item["servicetype@OData.Community.Display.V1.FormattedValue"],
    servicetype: item.servicetype,
    author: item.author,
    filesize: item.filesize
  }));
}

async function getSiteUrl(
  context: ComponentFramework.Context<IInputs>,
  parentLocationId: string
): Promise<string> {
  const metadata = await getEntityMetadata(context);
  const url = `${metadata?.clientUrl}/api/data/v9.0/FetchSiteUrl`;
  const body = JSON.stringify({
    "DocumentId": parentLocationId,
    "ParentEntityReference": {
        "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
        [`${metadata?.schemaName}id`]: `${metadata?.entityId}`
    }
});

const fetchOptions = {
  method: 'POST',
  headers: {
      'Content-Type': 'application/json'
  },
  body: body
};
try {
  const response = await fetch(url, fetchOptions);
  if (!response.ok) {
    throw new Error('Error getting SharePoint site: ' + response.statusText);
  }
  const responseData = await response.json();
  if (!responseData.SiteAndLocationUrl) {
    throw new Error('SiteAndLocationUrl not found in response');
  }
  return responseData.SiteAndLocationUrl;
} catch (error) {
  console.error('Error getting SharePoint site: ', error);
  throw error;
}
}

export async function createSharePointLocation(
  context: ComponentFramework.Context<IInputs>,
  locationName: string,
  relativePath: string,
  parentLocationId: string,
): Promise<string> {
  const metadata = await getEntityMetadata(context);
  const url = `${metadata?.clientUrl}/api/data/v9.0/AddOrEditLocation`;
  const absUrl = await getSiteUrl(context, parentLocationId);
  const body = JSON.stringify({
    "AbsUrl": absUrl + "/" + locationName,
    "DocumentId": "",
    "IsAddOrEditMode": true,
    "IsCreateFolder": true,
    "LocationName": locationName,
    "ParentEntityReference": {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
      [`${metadata?.schemaName}id`]: metadata?.entityId,
    },
    "ParentId": parentLocationId,
    "ParentType": "sharepointdocumentlocation",
    "RelativePath": locationName
  });

  const fetchOptions = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'OData-Version': '4.0',
      'OData-MaxVersion': '4.0',
      'Accept': 'application/json',
    },
    body: body
  };

  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw new Error('Failed to create a location: ' + response.statusText);
    }
    const responseData = await response.json();
    if (!responseData.LocationId) {
      throw new Error('LocationId not found in response');
    }
    return responseData.LocationId;
  } catch (error) {
    console.error('Error creating a location:', error);
    throw error;
  }
}

export async function createSharePointDocument(
  context: ComponentFramework.Context<IInputs>,
  fileName: string,
  dataURL: string,
  folderPath: string,
  locationId: string
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  if(!metadata){
    return
  }
  const isActivity = isActivityType(metadata.schemaName);
  const url = `${metadata?.clientUrl}/api/data/v9.0/UploadDocument`;
  const base64 = dataURL.split(',')[1];
  const body = JSON.stringify({
      "Content": base64,
      "Entity": {
          "@odata.type": "Microsoft.Dynamics.CRM.sharepointdocument",
          "locationid": locationId,
          "title": `${fileName}`
      },
      "OverwriteExisting": true,
      "ParentEntityReference": {
          "@odata.type": `Microsoft.Dynamics.CRM.${metadata.schemaName}`,
          [`${isActivity?"activity":metadata.schemaName}id`]: `${metadata?.entityId}`
      },
      "FolderPath": `${folderPath}`
  });

  const fetchOptions = {
      method: 'POST',
      headers: {
          'Content-Type': 'application/json'
      },
      body: body
  };

  try {
      const response = await fetch(url, fetchOptions);
      if (!response.ok) {
          throw new Error('Failed to upload document: ' + response.statusText);
      }
  } catch (error) {
      console.error('Error uploading document:', error);
      throw error;
  }
}

export async function deleteSharePointDocument(
  context: ComponentFramework.Context<IInputs>,
  sharepointDocumentId: string,
  documentId: number,
  fileType: string,
  rawlocationId: string,
  isDefaultLocation: boolean
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return;
  }
  let locationId = rawlocationId;
  if(!isDefaultLocation){
    locationId = "00000000-0000-0000-0000-000000000000";
  }
  const isActivity = isActivityType(metadata.schemaName);
  const url = `${metadata?.clientUrl}/api/data/v9.0/DeleteDocument`;
  const body = JSON.stringify({
    "Entities": [
      {
        "@odata.type": "Microsoft.Dynamics.CRM.sharepointdocument",
        "sharepointdocumentid": `{${sharepointDocumentId.toUpperCase()}}`,
        "documentid": documentId,
        "filetype": fileType,
        "locationid": locationId
      }
    ],
    "ParentEntityReference": {
      "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
      [`${isActivity?"activity":metadata.schemaName}id`]: `${metadata?.entityId}`
    }
  });

  const fetchOptions = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: body
  };

  try {
    const response = await fetch(url, fetchOptions);
    if (!response.ok) {
      throw new Error('Failed to delete document: ' + response.statusText);
    }
  } catch (error) {
    console.error('Error deleting document:', error);
    throw error;
  }
}

export async function createSharePointFolder(
  context: ComponentFramework.Context<IInputs>,
  folderName: string,
  folderPath: string,
  rawlocationId: string,
  isDefaultLocation: boolean
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  if(!metadata){
    return
  }

  let locationId = rawlocationId;
  if(!isDefaultLocation){
    locationId = "00000000-0000-0000-0000-000000000000";
  }
  
  const isActivity = isActivityType(metadata.schemaName);
  const url = `${metadata?.clientUrl}/api/data/v9.0/NewDocument`;
  const body = JSON.stringify({
      "FileName": folderName,
      "LocationId": locationId,
      "IsFolder": true,
      "FolderPath": folderPath,
      "ParentEntityReference": {
          "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
          [`${isActivity?"activity":metadata.schemaName}id`]: `${metadata?.entityId}`
      }
  });

  const fetchOptions = {
      method: 'POST',
      headers: {
          'Content-Type': 'application/json'
      },
      body: body
  };

  try {
      const response = await fetch(url, fetchOptions);
      if (!response.ok) {
          throw new Error('Failed to create a folder: ' + response.statusText);
      }
  } catch (error) {
      console.error('Error creating a folder:', error);
      throw error;
  }
}

export async function createAtivityDocument(
  context: ComponentFramework.Context<IInputs>,
  fileName: string,
  base64: string,
  mimeType: string
): Promise<GenericActionResponse> {
  const metadata = await getEntityMetadata(context);
  if(!metadata){
    return { success: false, message: `Unable to find entity metadata to create an activity document` };
  }
  const entity = "activitymimeattachment";
  const objectId = `objectid_${metadata.schemaName}@odata.bind`;
  const record = {
    "filename": fileName,
    "mimetype": mimeType,
    "body": base64.split(',')[1],
    [objectId]: `/${metadata?.logicalCollectionName}(${metadata.entityId})`,
    "objecttypecode": metadata.schemaName
  };

  try {
    await context.webAPI.createRecord(entity, record);
    return { success: true, message: `${fileName} added to ${metadata.schemaName} attachments successfully`};
  } catch (error) {
    console.error("Error adding SharePoint document to activity:", error);
    return { success: false, message: `Error adding SharePoint document to activity: ${(error as any).message}` };
  }
}

