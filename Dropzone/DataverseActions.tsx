import { IInputs } from "./generated/ManifestTypes";

export async function createRelatedNote(
  context: ComponentFramework.Context<IInputs>,
  fileName: string,
  base64: string,
  fileSize: number,
  mimeType: string
) {
  const entity = "annotation";
  const objectId = `objectid_${(context as any).page.entityTypeName}@odata.bind`;
  const record = {
    "isdocument": true,
    "filename": fileName,
    "subject": fileName,
    "filesize": fileSize,
    "mimetype": mimeType,
    "documentbody": base64.split(',')[1],
    [objectId]: `/${(context as any).page.entityTypeName}s(${(context as any).page.entityId})`
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
  const entityName = (context as any).page.entityTypeName;
  const entityId = (context as any).page.entityId;
  const filter = `_objectid_value eq ${entityId}`;

  try {
    const result = await context.webAPI.retrieveMultipleRecords('annotation', `?$filter=${filter}&$orderby=createdon desc`);
    const filteredEntities = result.entities.filter((entity: any) => entity.objecttypecode === entityName);
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
