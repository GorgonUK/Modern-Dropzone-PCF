import React, { Component } from "react";
import Dropzone from "react-dropzone";
import { FileIcon, defaultStyles, DefaultExtensionType } from "react-file-icon";
import { IInputs } from "../generated/ManifestTypes";
import {
  createRelatedNote,
  getRelatedNotes,
  updateRelatedNote,
  deleteRelatedNote,
  duplicateRelatedNote,
  getSharePointLocations,
  getSharePointData,
  getSharePointFolderData,
} from "../DataverseActions";
import { FileData, SharePointDocument } from "../Interfaces";
import {
  DefaultButton,
  PrimaryButton,
  TextField,
  Dialog,
  DialogType,
  DialogFooter,
  Stack,
  IStackStyles,
  IStackTokens,
  CommandButton,
  SearchBox,
  Spinner,
  SpinnerSize,
  Toggle,
  TooltipHost,
  IButtonStyles,
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
} from "@fluentui/react";
import { Tooltip } from "react-tippy";
import "react-tippy/dist/tippy.css";
import toast, { Toaster } from "react-hot-toast";
import FolderIcon from "./FolderIcon";
import "../css/Dropzone.css";

export interface LandingProps {
  context?: ComponentFramework.Context<IInputs>;
}

interface LandingState {
  files: FileData[];
  editingFileId?: string;
  selectedFiles: string[];
  searchText: string;
  sortAsc: boolean;
  isLoading: boolean;
  sharePointDocLoc: boolean;
  showTooltip: boolean;
  documentLocations: { name: string; sharepointdocumentlocationid: string }[];
  selectedDocumentLocation: string | null;
  sharePointData: SharePointDocument[];
  currentFolderPath: string;
  folderStack: string[];
}

type FileAction = "edit" | "download" | "duplicate" | "delete";

const SharePointDocLocTooltip = (
  <span>
    Upload files directly to SharePoint. To learn more go to{" "}
    <a
      href="https://learn.microsoft.com/en-us/power-platform/admin/create-edit-document-location-records"
      target="_blank"
    >
      MS Docs
    </a>
    .
  </span>
);

const tooltipId = "sharePointDocLocTooltip";
const buttonStyles: IButtonStyles = {
  root: {
    border: "none",
    minWidth: "auto",
    margin: 0,
    padding: 0,
  },
  rootHovered: {
    border: "none",
  },
  rootPressed: {
    border: "none",
  },
  rootExpanded: {
    border: "none",
  },
  rootChecked: {
    border: "none",
  },
  rootDisabled: {
    border: "none",
  },
};
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};
const ribbonStyles: IStackStyles = {
  root: {
    alignItems: "center",
    display: "flex",
    width: "100%",
    justifyContent: "space-between",
  },
};

const searchRibbonStyles: IStackStyles = {
  root: {
    marginBottom: "0px",
  },
};

const isValidBase64 = (str: string) => {
  try {
    return btoa(atob(str)) === str;
  } catch (err) {
    return false;
  }
};

const ribbonStackTokens: IStackTokens = { childrenGap: 10 };

export class Landing extends Component<LandingProps, LandingState> {
  constructor(props: LandingProps) {
    super(props);
    this.removeFile = this.removeFile.bind(this);
    this.downloadFile = this.downloadFile.bind(this);
    this.state = {
      files: [],
      editingFileId: undefined,
      selectedFiles: [],
      searchText: "",
      sortAsc: true,
      isLoading: true,
      sharePointDocLoc: false,
      showTooltip: false,
      documentLocations: [],
      selectedDocumentLocation: null,
      sharePointData: [],
      currentFolderPath: "",
      folderStack: [],
    };
    this.toggleSharePointDocLoc = this.toggleSharePointDocLoc.bind(this);
    this.toggleTooltip = this.toggleTooltip.bind(this);
    this.getSharePointLocations = this.getSharePointLocations.bind(this);
    this.getSharePointData = this.getSharePointData.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleBackClick = this.handleBackClick.bind(this);
  }
  private async handleFolderClick(folderPath: string): Promise<void> {
    this.setState(
      (prevState) => ({
        folderStack: [...prevState.folderStack, prevState.currentFolderPath],
      }),
      async () => {
        try {
          const data = await getSharePointFolderData(
            this.props.context!,
            folderPath
          );
          this.setState({
            sharePointData: data,
            currentFolderPath: folderPath,
          });
        } catch (error) {
          console.error("Error fetching folder data:", error);
        }
      }
    );
  }
  private handleBackClick = async () => {
    this.setState(
      (prevState) => {
        const folderStack = [...prevState.folderStack];
        const previousFolderPath = folderStack.pop();
        return { folderStack, currentFolderPath: previousFolderPath || "" };
      },
      async () => {
        const { currentFolderPath } = this.state;
        const data = await getSharePointFolderData(
          this.props.context!,
          currentFolderPath
        );
        this.setState({ sharePointData: data });
      }
    );
  };

  private async getSharePointLocations(): Promise<void> {
    const { context } = this.props;
    try {
      const response = await getSharePointLocations(context!);
      const firstLocation =
        response.length > 0 ? response[0].sharepointdocumentlocationid : null;
      this.setState({
        documentLocations: response,
        selectedDocumentLocation: firstLocation,
      });
    } catch (error) {
      console.error("Error fetching SharePoint document locations:", error);
    }
  }
  private async getSharePointData(): Promise<void> {
    this.setState(
      (prevState) => {
        const folderStack = [...prevState.folderStack];
        const previousFolderPath = folderStack.pop();
        return { folderStack, currentFolderPath: previousFolderPath || "" };
      },
      async () => {
        const { currentFolderPath } = this.state;
        const data = await getSharePointFolderData(
          this.props.context!,
          currentFolderPath
        );
        this.setState({ sharePointData: data });
      }
    );
  }
  componentDidMount() {
    this.loadExistingFiles().then(() => {
      this.setState({ isLoading: false });
    });
  }

  getFileExtension(filename: string): DefaultExtensionType {
    const extension = filename.split(".").pop()?.toLowerCase() as
      | DefaultExtensionType
      | undefined;
    return extension && extension in defaultStyles ? extension : "txt";
  }

  handleDrop = (acceptedFiles: File[]) => {
    acceptedFiles.forEach((file) => {
      const reader = new FileReader();
// add condition for SharePoint file upload
      reader.onload = () => {
        const binaryStr = reader.result as string;
        const createNotePromise = createRelatedNote(
          this.props.context!,
          file.name,
          binaryStr,
          file.size,
          file.type
        );

        toast.promise(createNotePromise, {
          loading: "Uploading file...",
          success: (res) => {
            if (res.noteId) {
              this.setState((prevState) => {
                const newFiles = [
                  ...prevState.files,
                  {
                    filename: file.name,
                    filesize: file.size,
                    documentbody: binaryStr,
                    mimetype: file.type,
                    noteId: res.noteId,
                    createdon: new Date(),
                    subject: "",
                    notetext: "",
                  },
                ];
                newFiles.sort(
                  (a, b) => b.createdon.getTime() - a.createdon.getTime()
                );
                return { files: newFiles };
              });
              return `File ${file.name} uploaded successfully!`;
            } else {
              throw new Error("Note ID was not returned");
            }
          },
          error: "Error uploading file",
        });
      };

      reader.onerror = () => {
        toast.error(`Error reading file: ${file.name}`);
      };

      reader.readAsDataURL(file);
    });
  };

  loadExistingFiles = async () => {
    if (!this.props.context) {
      console.error("Component Framework context is not available.");
      return;
    }
    this.setState({ isLoading: true });
    const { sharePointDocLoc } = this.state;
    if (!sharePointDocLoc) {
      const response = await getRelatedNotes(this.props.context);
      if (response.success) {
        const filesData: FileData[] = response.data.map((item: any) => ({
          filename: item.filename,
          filesize: item.filesize,
          documentbody: item.documentbody,
          mimetype: item.mimetype,
          noteId: item.annotationid,
          createdon: new Date(item.createdon),
          subject: item.subject,
          notetext: item.notetext,
        }));
        this.setState({ files: filesData });
      } else {
        console.error("Failed to retrieve files:", response.message);
      }
    } else {
      this.getSharePointLocations();
      this.getSharePointData();
    }
    this.setState({ isLoading: false });
  };

  downloadFile = (fileData: FileData) => {
    if (!fileData.documentbody || !fileData.mimetype || !fileData.filename) {
      toast.error("Missing file data for download");
      return;
    }

    try {
      toast.loading("Preparing download...");
      let base64Data = fileData.documentbody;

      const dataUrlPrefix = "base64,";
      if (base64Data.includes(dataUrlPrefix)) {
        const base64Index =
          base64Data.indexOf(dataUrlPrefix) + dataUrlPrefix.length;
        base64Data = base64Data.substring(base64Index);
      }

      if (!isValidBase64(base64Data)) {
        throw new Error("Invalid base64 string");
      }
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: fileData.mimetype });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.setAttribute("download", fileData.filename);
      document.body.appendChild(link);
      link.click();
      if (link.parentNode) {
        link.parentNode.removeChild(link);
      }
      URL.revokeObjectURL(url);
      toast.dismiss();
      toast.success(`${fileData.filename} downloaded successfully!`);
    } catch (error) {
      toast.error(`Failed to download file: ${(error as Error).message}`);
    }
  };

  duplicateFile = async (noteId: string) => {
    const { context } = this.props;

    const duplicationPromise = duplicateRelatedNote(context!, noteId);
    toast.promise(duplicationPromise, {
      loading: "Duplicating file...",
      success: (response) => {
        this.loadExistingFiles();
        return `File duplicated successfully!`;
      },
      error: (err) => `Failed to duplicate file: ${err.message}`,
    });
  };

  removeFile = async (noteId?: string) => {
    if (!noteId) {
      console.error("No note ID provided for deletion");
      toast.error("No note ID provided");
      return;
    }

    this.setState((prevState) => ({
      files: prevState.files.map((file) =>
        file.noteId === noteId ? { ...file, isLoading: true } : file
      ),
    }));

    toast.promise(deleteRelatedNote(this.props.context!, noteId), {
      loading: "Deleting file...",
      success: () => {
        this.setState((prevState) => ({
          files: prevState.files.filter((file) => file.noteId !== noteId),
        }));
        return "File successfully deleted!";
      },
      error: (err) => {
        this.setState((prevState) => ({
          files: prevState.files.map((file) =>
            file.noteId === noteId ? { ...file, isLoading: false } : file
          ),
        }));
        console.error("Toast Error: ", err);
        return `Failed to delete file: ${err.message || "Unknown error"}`;
      },
    });
  };

  toggleEditModal = (noteId?: string) => {
    this.setState({ editingFileId: noteId });
  };

  saveChanges = async (noteId: string, filename: string): Promise<void> => {
    toast
      .promise(updateRelatedNote(this.props.context!, noteId, filename), {
        loading: "Saving changes...",
        success: "Changes saved successfully!",
        error: "Failed to save changes",
      })
      .then((response) => {
        if (response.success) {
          this.setState((prevState) => ({
            files: prevState.files.map((file) =>
              file.noteId === noteId
                ? { ...file, filename, isEditing: false }
                : file
            ),
          }));
          this.toggleEditModal();
        } else {
          console.error("Failed to update note:", response.message);
        }
      });
  };

  formatFileSize(sizeInBytes: number) {
    const sizeInKB = sizeInBytes / 1024;
    const sizeInMB = sizeInBytes / 1048576;

    if (sizeInMB >= 1) {
      return `${sizeInMB.toFixed(2)} MB`;
    } else if (sizeInKB >= 1) {
      return `${sizeInKB.toFixed(2)} KB`;
    } else {
      return `${sizeInBytes} Bytes`;
    }
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

  performActionOnSelectedFiles = (action: FileAction) => {
    this.state.selectedFiles.forEach((noteId) => {
      const file = this.state.files.find((f) => f.noteId === noteId);
      if (!file) {
        toast.error("File not found.");
        return;
      }

      if (action === "duplicate") {
        this.duplicateFile(noteId);
      } else if (action === "download") {
        this.downloadFile(file);
      } else if (action === "delete") {
        this.removeFile(noteId);
      } else if (action === "edit") {
        this.toggleEditModal(noteId);
      }
    });

    this.setState({ selectedFiles: [] });
  };
  handleSearch = (event: any, newValue?: string) => {
    this.setState({ searchText: newValue || "" });
  };

  toggleSortOrder = () => {
    this.setState((prevState) => ({ sortAsc: !prevState.sortAsc }));
  };
  toggleTooltip() {
    this.setState((prevState) => ({ showTooltip: !prevState.showTooltip }));
  }

  toggleSharePointDocLoc(_: React.MouseEvent<HTMLElement>, checked?: boolean) {
    this.setState({ sharePointDocLoc: !!checked }, () => {
      this.loadExistingFiles();
    });
  }
  getFilteredAndSortedFiles() {
    const { files, searchText } = this.state;
    return files.filter(file => file.filename.toLowerCase().includes(searchText.toLowerCase()));
  }
  
  getFilteredAndSortedSPFiles() {
    const { sharePointData, searchText } = this.state;
    return sharePointData.filter(spData => spData.fullname.toLowerCase().includes(searchText.toLowerCase()));
  }

  private handleDropdownChange(
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (option) {
      this.setState({ selectedDocumentLocation: option.key as string });
    }
  }

  renderRibbon = () => {
    const { selectedFiles } = this.state;
    return (
      <>
        <Stack
          horizontal
          styles={searchRibbonStyles}
          className="easeIn wideRibbon"
        >
          <Stack horizontal styles={ribbonStyles} className="easeIn">
            <Stack horizontal tokens={ribbonStackTokens}>
              <SearchBox
                placeholder="Search files..."
                onSearch={(newValue) => this.handleSearch(null, newValue)}
                onChange={(_, newValue) => this.handleSearch(null, newValue)}
                styles={{
                  root: {
                    width: 200,
                    border: "none",
                    boxShadow: "none",
                    position: "relative",
                    selectors: {
                      ":hover": {
                        border: "none",
                        boxShadow: "none",
                      },
                      ":focus": {
                        border: "none",
                        boxShadow: "none",
                      },
                      "::after": {
                        content: "none !important",
                      },
                    },
                  },
                  field: {
                    border: "none",
                    boxShadow: "none",
                  },
                  iconContainer: {
                    color: "rgb(17, 94, 163)",
                    selectors: {
                      ":hover": {
                        color: "rgb(17, 94, 163)",
                      },
                      ":focus": {
                        color: "rgb(17, 94, 163)",
                      },
                    },
                  },
                }}
              />
            </Stack>
            {selectedFiles.length > 0 && (
              <Stack horizontal tokens={ribbonStackTokens}>
                {selectedFiles.length < 2 && (
                  <CommandButton
                    iconProps={{ iconName: "Edit" }}
                    text="Rename"
                    onClick={() => this.performActionOnSelectedFiles("edit")}
                    disabled={selectedFiles.length === 0}
                    className="icon-button"
                  />
                )}
                <CommandButton
                  iconProps={{ iconName: "Download" }}
                  text="Download"
                  onClick={() => this.performActionOnSelectedFiles("download")}
                  disabled={selectedFiles.length === 0}
                  className="icon-button"
                />
                <CommandButton
                  iconProps={{ iconName: "Copy" }}
                  text="Duplicate"
                  onClick={() => this.performActionOnSelectedFiles("duplicate")}
                  disabled={selectedFiles.length === 0}
                  className="icon-button"
                />
                <CommandButton
                  iconProps={{ iconName: "Delete" }}
                  text="Delete"
                  onClick={() => this.performActionOnSelectedFiles("delete")}
                  disabled={selectedFiles.length === 0}
                  className="icon-button"
                />
              </Stack>
            )}
          </Stack>
        </Stack>
      </>
    );
  };

  toggleFileSelection = (noteId: string, file?: SharePointDocument) => {
    if (file && file.filetype === "folder") {
      this.handleFolderClick(file.relativelocation);
      return;
    }

    const isSelected = this.state.selectedFiles.includes(noteId);
    this.setState((prevState) => ({
      selectedFiles: isSelected
        ? prevState.selectedFiles.filter((id) => id !== noteId)
        : [...prevState.selectedFiles, noteId],
    }));
  };
  renderFileList() {
    const { sharePointDocLoc, selectedFiles, currentFolderPath } = this.state;
    const files = this.getFilteredAndSortedFiles();
    const sharePointData = this.getFilteredAndSortedSPFiles();
    if (sharePointDocLoc) {
      return (
        <>
          {currentFolderPath && (
            <button
              onClick={(event) => {
                event.preventDefault();
                event.stopPropagation();
                this.handleBackClick();
              }}
              className="back-button"
            >
              Back
            </button>
          )}
          {sharePointDocLoc
            ? sharePointData.map((file) => (
                <div
                  key={file.sharepointdocumentid}
                  className={`file-box ${
                    selectedFiles.includes(file.sharepointdocumentid)
                      ? "selected"
                      : ""
                  }`}
                  onClick={(event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    this.toggleFileSelection(file.sharepointdocumentid, file);
                  }}
                >
                  <div className="file-image">
                    {file.filetype === "folder" ? (
                      <FolderIcon />
                    ) : (
                      <FileIcon
                        extension={this.getFileExtension(file.fullname)}
                        {...defaultStyles[this.getFileExtension(file.fullname)]}
                      />
                    )}
                  </div>
                  <Tooltip
                    title={file.fullname}
                    position="top"
                    trigger="mouseenter"
                    arrow={true}
                    arrowSize="regular"
                    theme="light"
                  >
                    <p className="file-name">
                      {this.middleEllipsis(file.fullname)}
                    </p>
                  </Tooltip>

                  {file.filetype !== "folder" ? (
                    <p className="file-size">
                      {this.formatFileSize(file.filesize)}
                    </p>
                  ) : (
                    <p className="file-size"> </p>
                  )}
                </div>
              ))
            : null}
        </>
      );
    } else {
      return files.map((file) => (
        <div
          key={file.noteId}
          className={`file-box ${
            selectedFiles.includes(file.noteId || "") ? "selected" : ""
          }`}
          onClick={(event) => {
            event.preventDefault();
            event.stopPropagation();
            this.toggleFileSelection(file.noteId!);
          }}
        >
          <div className="file-image">
            <FileIcon
              extension={this.getFileExtension(file.filename)}
              {...defaultStyles[this.getFileExtension(file.filename)]}
            />
          </div>
          <Tooltip
            title={file.filename}
            position="top"
            trigger="mouseenter"
            arrow={true}
            arrowSize="regular"
            theme="light"
          >
            <p className="file-name">{this.middleEllipsis(file.filename)}</p>
          </Tooltip>
          <p className="file-size">{this.formatFileSize(file.filesize)}</p>
        </div>
      ));
    }
  }

  render() {
    const {
      editingFileId,
      isLoading,
      sharePointDocLoc,
      showTooltip,
      selectedDocumentLocation,
      documentLocations,
      currentFolderPath,
    } = this.state;
    const dropdownOptions: IDropdownOption[] = documentLocations.map(
      (location) => ({
        key: location.sharepointdocumentlocationid,
        text: location.name,
      })
    );
    const files = this.getFilteredAndSortedFiles();
    const isEmpty = files.length === 0;
    const entityIdExists = (this.props.context as any).page.entityId;
    const editingFile = files.find((file) => file.noteId === editingFileId);
    const docLoctooltip = (
      <TooltipHost
        content={showTooltip ? SharePointDocLocTooltip : undefined}
        id={tooltipId}
      >
        <DefaultButton
          aria-label="more info"
          aria-describedby={showTooltip ? tooltipId : undefined}
          onClick={this.toggleTooltip}
          styles={buttonStyles}
          iconProps={{ iconName: "Info" }}
        />
      </TooltipHost>
    );

    if (!entityIdExists) {
      return (
        <div className="record-not-created-message">
          This record hasn&apos;t been created yet. To enable file upload,
          create this record.
        </div>
      );
    }
    return (
      <>
        <Toaster position="top-right" reverseOrder={false} />

        {editingFile && (
          <Dialog
            hidden={!editingFileId}
            onDismiss={() => this.toggleEditModal()}
            dialogContentProps={{
              type: DialogType.normal,
              title: "Edit file name",
            }}
          >
            <TextField
              label="File name"
              value={editingFile.filename || ""}
              onChange={(e, newValue) => {
                const updatedFiles = files.map((file) =>
                  file.noteId === editingFileId
                    ? { ...file, filename: newValue || "" }
                    : file
                );
                this.setState({ files: updatedFiles });
              }}
            />
            <DialogFooter>
              <PrimaryButton
                onClick={() => {
                  this.saveChanges(editingFile.noteId!, editingFile.filename!);
                }}
                text="Save"
              />
              <DefaultButton
                onClick={() => this.toggleEditModal()}
                text="Cancel"
              />
            </DialogFooter>
          </Dialog>
        )}
        <div className="ribbon-dropzone-wrapper">
          {this.renderRibbon()}

          <Dropzone onDrop={this.handleDrop}>
            {({ getRootProps, getInputProps }) => (
              <div className="dropzone-wrapper">
                <div
                  {...getRootProps()}
                  className={`dropzone ${isEmpty ? "empty" : ""}`}
                >
                  {isLoading ? (
                    <div className="spinner-box">
                      <Spinner size={SpinnerSize.medium} />
                    </div>
                  ) : (
                    <>
                      <input {...getInputProps()} />
                      {isEmpty ? (
                        <p>Drag and drop files here or Browse for files</p>
                      ) : (
                        this.renderFileList()
                      )}
                    </>
                  )}
                </div>
              </div>
            )}
          </Dropzone>
          <Toggle
            label={<div>SharePoint Documents {docLoctooltip}</div>}
            inlineLabel
            onText="On"
            offText="Off"
            checked={sharePointDocLoc}
            onChange={this.toggleSharePointDocLoc}
          />
          {sharePointDocLoc && (
            <Dropdown
              options={dropdownOptions}
              disabled={false}
              styles={dropdownStyles}
              onChange={this.handleDropdownChange}
              defaultSelectedKey={selectedDocumentLocation}
            />
          )}
        </div>
      </>
    );
  }
}

export default Landing;
