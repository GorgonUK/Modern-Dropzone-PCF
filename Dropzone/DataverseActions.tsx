import { IInputs } from "./generated/ManifestTypes";
import { getEntityMetadata } from "./utils";

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

export async function duplicateRelatedNote(
  context: ComponentFramework.Context<IInputs>,
  noteId: string
) {
  const entity = "annotation";
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return { success: false, message: "Failed to retrieve entity metadata" };
  }
  const objectId = `objectid_${metadata.schemaName}@odata.bind`;
  try {
    const originalNote = await context.webAPI.retrieveRecord(entity, noteId);
    const newNoteData = {
      "isdocument": originalNote.isdocument,
      "filename": originalNote.filename,
      "subject": originalNote.subject,
      "notetext": originalNote.notetext,
      "filesize": originalNote.filesize,
      "mimetype": originalNote.mimetype,
      "documentbody": originalNote.documentbody,
      [objectId]: `/${metadata.logicalCollectionName}(${metadata.entityId})`
    };

    const creationResult = await context.webAPI.createRecord(entity, newNoteData);
    return { success: true, message: "Note duplicated successfully", newNoteId: creationResult.id };
  } catch (error) {
    console.error("Error duplicating note:", error);
    return { success: false, message: `Error duplicating note: ${(error as any).message}` };
  }
}

export async function getSharePointLocations(context: ComponentFramework.Context<IInputs>): Promise<{ name: string, sharepointdocumentlocationid:string }[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  const fetchXml = `
  <fetch mapping="logical" version="1.0">
  <entity name="sharepointdocumentlocation">
    <attribute name="name" />
    <attribute name="sharepointdocumentlocationid" />
    <filter type="and">
      <condition attribute="regardingobjectid" operator="eq" value="${metadata.entityId}" />
      <condition attribute="statecode" operator="eq" value="0" />
      <condition attribute="statuscode" operator="eq" value="1" />
      <condition attribute="sitecollectionid" operator="not-null" />
      <condition attribute="absoluteurl" operator="null" />
    </filter>
    <filter type="or">
      <filter type="and">
        <condition attribute="servicetype" operator="eq" value="0" />
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

  const data = await response.json();

  return data.value.map((item: any) => ({
    name: item.name,
    sharepointdocumentlocationid: item.sharepointdocumentlocationid
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
        <link-entity name="${metadata.schemaName}" from="${metadata.schemaName}id" to="regardingobjectid" alias="bb">
          <filter type="and">
            <condition attribute="${metadata.schemaName}id" operator="eq" uitype="${metadata.schemaName}" value="${metadata.entityId}"/>
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

export async function getSharePointFolderData(context: ComponentFramework.Context<IInputs>,folderPath?: string,selectedDocumentLocation?:string,selectedDocumentLocationName?:string): Promise<any[]> {
  const metadata = await getEntityMetadata(context);
  if (!metadata) {
    return [];
  }
  let folderFilter = '';
  let filters = '';
  if (folderPath) {
    folderFilter = `
      <filter type="and">
        <condition attribute="relativelocation" operator="eq" value="${folderPath}"/>
      </filter>`;
  }
  if(selectedDocumentLocation){
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
        <link-entity name="${metadata.schemaName}" from="${metadata.schemaName}id" to="regardingobjectid" alias="bb">
          <filter type="and">
            <condition attribute="${metadata.schemaName}id" operator="eq" uitype="${metadata.schemaName}" value="${metadata.entityId}"/>
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

export async function createSharePointDocument(
  context: ComponentFramework.Context<IInputs>,
  fileName: string,
  dataURL: string,
  folderPath: string,
  locationId: string
): Promise<void> {
  const metadata = await getEntityMetadata(context);
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
          "@odata.type": `Microsoft.Dynamics.CRM.${metadata?.schemaName}`,
          [`${metadata?.schemaName}id`]: `${metadata?.entityId}`
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
  locationId: string
): Promise<void> {
  const metadata = await getEntityMetadata(context);
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
  locationId: string
): Promise<void> {
  const metadata = await getEntityMetadata(context);
  const url = `${metadata?.clientUrl}/api/data/v9.0/NewDocument`;
  const body = JSON.stringify({
      "FileName": folderName,
      "LocationId": locationId,
      "IsFolder": true,
      "FolderPath": folderPath,
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
          throw new Error('Failed to create a folder: ' + response.statusText);
      }
  } catch (error) {
      console.error('Error creating a folder:', error);
      throw error;
  }
}

