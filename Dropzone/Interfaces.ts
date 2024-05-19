export interface FileData {
    filename: string;
    filesize: number;
    documentbody?: string;
    mimetype?: string;
    noteId?: string;
    createdon: Date;
    isLoading?: boolean;
  }