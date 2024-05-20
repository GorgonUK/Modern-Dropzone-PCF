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
    console.log(result)
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

export async function updateRelatedNote(
  context: ComponentFramework.Context<IInputs>,
  noteId: string,
  subject: string,
  notetext: string
): Promise<{ success: boolean; message: string }> {
  const data = {
    subject: subject,
    notetext: notetext
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
  const objectId = `objectid_${(context as any).page.entityTypeName}@odata.bind`;
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
      [objectId]: `/${(context as any).page.entityTypeName}s(${(context as any).page.entityId})`
    };

    const creationResult = await context.webAPI.createRecord(entity, newNoteData);
    console.log("Note duplicated successfully:", creationResult);
    return { success: true, message: "Note duplicated successfully", newNoteId: creationResult.id };
  } catch (error) {
    console.error("Error duplicating note:", error);
    return { success: false, message: `Error duplicating note: ${(error as any).message}` };
  }
}

