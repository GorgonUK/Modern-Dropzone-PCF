import React, { Component } from "react";
import Dropzone from "react-dropzone";
import { FileIcon, defaultStyles, DefaultExtensionType } from "react-file-icon";
import { IInputs } from "../generated/ManifestTypes";
import { createRelatedNote, getRelatedNotes, deleteRelatedNote } from "../DataverseActions";
import { FileData } from "../Interfaces";
import { v4 as uuidv4 } from "uuid";
import "./Landing.css"
import { IconButton } from '@fluentui/react';

export interface LandingProps {
  context?: ComponentFramework.Context<IInputs>;
}

interface LandingState {
  files: FileData[];
}

export class Landing extends Component<LandingProps, LandingState> {
  constructor(props: LandingProps) {
    super(props);
    this.removeFile = this.removeFile.bind(this);
    this.downloadFile = this.downloadFile.bind(this);
    this.state = {
      files: []
    };
  }

  componentDidMount() {
    this.loadExistingFiles();
  }

  getFileExtension(filename: string): DefaultExtensionType {
    const extension = filename.split(".").pop()?.toLowerCase() as DefaultExtensionType | undefined;
    return extension && extension in defaultStyles ? extension : "txt";
  }

  handleDrop = (acceptedFiles: File[]) => {
    acceptedFiles.forEach((file) => {
      const reader = new FileReader();
      reader.onload = async () => {
        const binaryStr = reader.result as string;
        const response = await createRelatedNote(
          this.props.context!,
          file.name,
          binaryStr,
          file.size,
          file.type
        );
        if (response.success && response.noteId) {
          this.setState(prevState => {
            const newFiles = [
              ...prevState.files,
              {
                filename: file.name,
                filesize: file.size,
                documentbody: binaryStr,
                mimetype: file.type,
                noteId: response.noteId,
                createdon: new Date()
              }
            ];
            newFiles.sort((a, b) => b.createdon.getTime() - a.createdon.getTime());
            return { files: newFiles };
          });
        } else {
          console.error(response.message);
        }
      };
      reader.readAsDataURL(file);
    });
  };

  loadExistingFiles = async () => {
    if (!this.props.context) {
      console.error("Component Framework context is not available.");
      return;
    }
    const response = await getRelatedNotes(this.props.context);
    if (response.success) {
      const filesData: FileData[] = response.data.map((item: any) => ({
        filename: item.filename,
        filesize: item.filesize,
        documentbody: item.documentbody,
        mimetype: item.mimetype,
        noteId: item.annotationid,
        createdon: new Date(item.createdon)
      }));
      this.setState({ files: filesData });
    } else {
      console.error("Failed to retrieve files:", response.message);
    }
  };

  downloadFile = (fileData: FileData) => {
    if (!fileData.documentbody || !fileData.mimetype || !fileData.filename) {
      console.error('Missing file data for download');
      return;
    }

    const byteCharacters = atob(fileData.documentbody);
    const byteNumbers = new Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i++) {
      byteNumbers[i] = byteCharacters.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: fileData.mimetype });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', fileData.filename);
    document.body.appendChild(link);
    link.click();
    if (link.parentNode) {
      link.parentNode.removeChild(link);
    }
  };


  removeFile = async (noteId?: string) => {
    if (noteId) {
      this.setState(prevState => ({
        files: prevState.files.map(file =>
          file.noteId === noteId ? { ...file, isLoading: true } : file
        )
      }));
      const response = await deleteRelatedNote(this.props.context!, noteId);
      if (response.success) {
        this.setState(prevState => ({
          files: prevState.files.filter(file => file.noteId !== noteId)
        }));
      } else {
        console.error("Failed to delete file:", response.message);
        this.setState(prevState => ({
          files: prevState.files.map(file =>
            file.noteId === noteId ? { ...file, isLoading: false } : file
          )
        }));
      }
    }
  };

  formatFileSize(sizeInKB: number) {
    const sizeInMB = sizeInKB / 1024;
    return `${sizeInMB.toFixed(2)} MB`;
  }

  middleEllipsis(filename: string, maxLength: number = 19): string {
    if (filename.length < maxLength) {
      return filename;
    }
    const lastDotIndex = filename.lastIndexOf(".");
    const extension = filename.substring(lastDotIndex + 1);
    const name = filename.substring(0, lastDotIndex);
    const startChars = 9;
    const endChars = 3;
    const start = name.substring(0, startChars);
    const end = name.substring(name.length - endChars);

    if (end.length < 3 && name.length - 3 > startChars) {
      return `${start}...${name.slice(-3)}.${extension}`;
    }
    return `${start}...${end}.${extension}`;
  }

  render() {
    const { files } = this.state;
    const isEmpty = files.length === 0;
    const entityIdExists = (this.props.context as any).page.entityId;

    if (!entityIdExists) {
      return (
        <div className="record-not-created-message">
          This record hasn&apos;t been created yet. To enable file upload, create this record.
        </div>
      );
    }
    return (
      <Dropzone onDrop={this.handleDrop}>
        {({ getRootProps, getInputProps }) => (
          <div className="dropzone-wrapper">
            <div {...getRootProps()} className={`dropzone ${isEmpty ? "empty" : ""}`}>
              <input {...getInputProps()} />
              {isEmpty ? (
                <p>Drag &apos;n&apos; drop files here or click to select files</p>
              ) : (
                files.map(file => (
                  <div key={file.noteId || uuidv4()} className="file-box" onClick={(event) => {
                    event.stopPropagation();
                  }}>
                    <div className="file-image"><FileIcon extension={this.getFileExtension(file.filename)} {...defaultStyles[this.getFileExtension(file.filename)]} /></div>
                    <p className="file-name">{this.middleEllipsis(file.filename)}</p>
                    <p className="file-size">{this.formatFileSize(file.filesize)}</p>
                    <div className="action-icons">
                      <IconButton
                        iconProps={{ iconName: 'Download' }}
                        title="Download"
                        ariaLabel="Download file"
                        onClick={(event) => {
                          event.preventDefault();
                          event.stopPropagation();
                          this.downloadFile(file);
                        }}
                        className="icon-button"
                      />
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Remove"
                        ariaLabel="Remove file"
                        onClick={(event) => {
                          event.preventDefault();
                          event.stopPropagation();
                          if (file.noteId) {
                            this.removeFile(file.noteId);
                          } else {
                            console.error('No note ID');
                          }
                        }}
                        disabled={file.isLoading}
                        className="icon-button"
                      />
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        )}
      </Dropzone>
    );
  }
}

export default Landing;
