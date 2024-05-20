export interface FileData {
  subject?: string;
  notetext?: string;
  filename: string;
  filesize: number;
  documentbody?: string;
  mimetype?: string;
  noteId?: string;
  createdon: Date;
  isLoading?: boolean;
  isEditing?: boolean;
}
