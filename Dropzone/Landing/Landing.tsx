import React, { Component } from "react";
import Dropzone from "react-dropzone";
import { IInputs } from "../generated/ManifestTypes";
import {
  createRelatedNote,
  getRelatedNotes,
  updateRelatedNote,
  deleteRelatedNote,
  duplicateRelatedNote,
  getSharePointLocations,
  getSharePointFolderData,
  createSharePointDocument,
  deleteSharePointDocument,
  createSharePointFolder,
} from "../DataverseActions";
import { FileData, SharePointDocument, PreviewFile } from "../Interfaces";
import { Img } from "react-image";
import {
  isPDF,
  isImage,
  isExcel,
  createDataUri,
} from "../utils";
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
  IconButton,
  Icon,
  ITooltipHostStyles,
  ContextualMenu,
  IContextualMenuProps,
  ContextualMenuItemType,
  Modal,
  IFocusTrapZoneProps,
} from "@fluentui/react";
import {
  getFileTypeIconProps,
  initializeFileTypeIcons
} from "@uifabric/file-type-icons";
import { read, utils } from "xlsx";
import Spreadsheet from "x-data-spreadsheet";
import 'x-data-spreadsheet/dist/xspreadsheet.css';
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
  spSearchText: string;
  notesSearchText: string;
  sortAsc: boolean;
  isLoading: boolean;
  sharePointDocLoc: boolean;
  showTooltip: boolean;
  documentLocations: { name: string; sharepointdocumentlocationid: string }[];
  selectedDocumentLocation: string | null;
  sharePointData: SharePointDocument[];
  currentFolderPath: string;
  folderStack: string[];
  showModal: boolean;
  newFolderName: string;
  selectedDocumentLocationName: string;
  isCollapsed: boolean;
  menuVisible: boolean;
  target: HTMLElement | null;
  previewFile: PreviewFile | null;
  isDialogOpen: boolean;
  xlsxContent: string;
  xlsxData: SheetData | null;
}

type FileAction = "edit" | "download" | "duplicate" | "delete" | "preview";

const SharePointDocLocTooltip = (
  <span>
    Upload files directly to a SharePoint Document Location. To learn more go to{" "}
    <a
      href="https://learn.microsoft.com/en-us/power-platform/admin/create-edit-document-location-records"
      target="_blank"
      rel="noopener noreferrer"
    >
      MS Docs
    </a>
    .
  </span>
);

const tooltipId = "sharePointDocLocTooltip";
const buttonStyles: Partial<ITooltipHostStyles> = {
  root: {
    border: "none",
    minWidth: "auto",
    margin: 0,
    padding: 0,
    display: "inline-block",
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

type SheetData = {
  [sheetName: string]: any[];
};

const isValidBase64 = (str: string) => {
  try {
    return btoa(atob(str)) === str;
  } catch (err) {
    return false;
  }
};

const ribbonStackTokens: IStackTokens = { childrenGap: 10 };

async function loadExcelFile(documentbody: string): Promise<SheetData> {
  const buffer = await fetch(documentbody).then(res => res.arrayBuffer());
  const workbook = read(buffer, { type: 'array' });
  const sheetsData: SheetData = {};

  workbook.SheetNames.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    sheetsData[sheetName] = utils.sheet_to_json(worksheet, { header: 1 });
  });

  return sheetsData;
}

export class Landing extends Component<LandingProps, LandingState> {
  constructor(props: LandingProps) {
    super(props);
    initializeFileTypeIcons();
    this.state = {
      files: [],
      editingFileId: undefined,
      selectedFiles: [],
      spSearchText: "",
      notesSearchText: "",
      sortAsc: true,
      isLoading: true,
      sharePointDocLoc: false,
      showTooltip: false,
      documentLocations: [],
      selectedDocumentLocation: null,
      sharePointData: [],
      currentFolderPath: "",
      folderStack: [],
      showModal: false,
      newFolderName: "",
      selectedDocumentLocationName: "",
      isCollapsed: window.innerWidth < 767,
      menuVisible: false,
      target: null,
      isDialogOpen: false,
      previewFile: null,
      xlsxContent: '',
      xlsxData: null
    };
    this.removeFile = this.removeFile.bind(this);
    this.downloadFile = this.downloadFile.bind(this);
    this.toggleSharePointDocLoc = this.toggleSharePointDocLoc.bind(this);
    this.toggleTooltip = this.toggleTooltip.bind(this);
    this.getSharePointLocations = this.getSharePointLocations.bind(this);
    this.getSharePointData = this.getSharePointData.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleFolderClick = this.handleFolderClick.bind(this);
    this.handleBackClick = this.handleBackClick.bind(this);
    this.handleDropdownChange = this.handleDropdownChange.bind(this);
    this.handleResize = this.handleResize.bind(this);
    this.toggleMenu = this.toggleMenu.bind(this);
    this.closeMenu = this.closeMenu.bind(this);
    this.openDialog = this.openDialog.bind(this);
    this.closeDialog = this.closeDialog.bind(this);
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
            folderPath,
            this.state.selectedDocumentLocation!,
            this.state.selectedDocumentLocationName
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
          currentFolderPath,
          this.state.selectedDocumentLocation!,
          this.state.selectedDocumentLocationName
        );
        this.setState({ sharePointData: data });
      }
    );
  };

  private async getSharePointLocations(): Promise<void> {
    const { context } = this.props;
    try {
      const response = await getSharePointLocations(context!);
      const firstLocationId =
        response.length > 0 ? response[0].sharepointdocumentlocationid : null;
      const firstLocationName = response.length > 0 ? response[0].name : null;
      this.setState({
        documentLocations: response,
        selectedDocumentLocation: firstLocationId,
        selectedDocumentLocationName: firstLocationName!,
      });
    } catch (error) {
      console.error("Error fetching SharePoint document locations:", error);
    }
  }

  private async getSharePointData(
    popFolderStack: boolean = false
  ): Promise<void> {
    this.setState(
      (prevState) => {
        const folderStack = [...prevState.folderStack];
        let currentFolderPath: string;

        if (popFolderStack) {
          currentFolderPath = folderStack.pop() || "";
        } else {
          currentFolderPath = prevState.currentFolderPath || "";
        }

        return { folderStack, currentFolderPath };
      },
      async () => {
        const { currentFolderPath } = this.state;
        const data = await getSharePointFolderData(
          this.props.context!,
          currentFolderPath,
          this.state.selectedDocumentLocation!,
          this.state.selectedDocumentLocationName
        );
        this.setState({ sharePointData: data });
      }
    );
  }

  componentDidMount() {
    this.loadExistingFiles().then(() => {
      this.setState({ isLoading: false });
    });
    window.addEventListener("resize", this.handleResize);
  }
  componentDidUpdate(prevProps: LandingProps, prevState: LandingState) {
   if (this.state.previewFile !== prevState.previewFile && this.state.previewFile && isExcel(this.state.previewFile.mimetype)) {
    const excelFile = this.state.previewFile.documentbody.replace(/(data:.*?;base64,).*?\1/, '$1');
    loadExcelFile(excelFile).then(data => {
      this.setState({ xlsxData: data });
      this.initializeSpreadsheet(data);
    });
    }
  }
  initializeSpreadsheet(data: SheetData) {
    const spreadsheet = new Spreadsheet('#xlsx-preview', {
      view: {
        height: () => document.documentElement.clientHeight - 40,
        width: () => document.documentElement.clientWidth - 60,
      },
      showToolbar: true,
      showGrid: true,
      showContextmenu: true,
    });
    const sheets = Object.keys(data).map((sheetName, index) => {
      const rows = data[sheetName].map((row: any, rowIndex: number) => {
        const cells = row.map((cell: any, colIndex: number) => ({
          text: cell,
          editable: false,
          className: 'readonly',
        }));
        return {
          cells,
          height: 20
        };
      });
      return {
        name: sheetName,
        rows,
        cols: data[sheetName][0].map((_: any, colIndex: number) => ({
          width: 100,
        })),
      };
    });
    spreadsheet.loadData(sheets);
  }
  componentWillUnmount() {
    window.removeEventListener("resize", this.handleResize);
  }

  handleResize() {
    this.setState({ isCollapsed: window.innerWidth < 864 });
  }

  toggleMenu(event: React.MouseEvent<HTMLElement>) {
    this.setState({
      menuVisible: !this.state.menuVisible,
      target: event.currentTarget,
    });
  }

  closeMenu() {
    this.setState({ menuVisible: false });
  }

  getFileExtension(filename?: string): string {
    if (!filename) {
      return "folder";
    }
    const extension = filename.split(".").pop()?.toLowerCase();
    return extension ? extension : "txt";
  }

  handleDrop = (acceptedFiles: File[]) => {
    if (!this.state.sharePointDocLoc) {
      acceptedFiles.forEach((file) => {
        const reader = new FileReader();
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
    } else {
      acceptedFiles.forEach((file) => {
        const reader = new FileReader();
        reader.onload = () => {
          const binaryStr = reader.result as string;
          const uploadPromise = createSharePointDocument(
            this.props.context!,
            file.name,
            binaryStr,
            this.state.currentFolderPath,
            this.state.selectedDocumentLocation!
          );

          toast.promise(uploadPromise, {
            loading: "Uploading to SharePoint...",
            success: () => {
              this.getSharePointData(false)
                .then(() => {
                  this.setState({ isLoading: false });
                })
                .catch((err) => {
                  console.error("Failed to reload files:", err);
                  this.setState({ isLoading: false });
                  toast.error("Failed to refresh file list.");
                });
              return `File ${file.name} uploaded successfully to SharePoint.`;
            },
            error: (err) => {
              this.setState({ isLoading: false });
              return `Error uploading to SharePoint: ${err.message}`;
            },
          });
        };

        reader.onerror = () => {
          toast.error(`Error reading file: ${file.name}`);
        };

        reader.readAsDataURL(file);
      });
    }
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
      if (this.state.documentLocations.length === 0) {
        this.getSharePointLocations().then(() => {
          this.getSharePointData();
        });
      } else if (this.state.documentLocations.length > 0) {
        this.getSharePointData();
      }
    }
    this.setState({ isLoading: false });
  };

  downloadFile = async (file: FileData | SharePointDocument) => {
    if (!this.state.sharePointDocLoc) {
      if (
        !("documentbody" in file) ||
        !file.documentbody ||
        !file.mimetype ||
        !file.filename
      ) {
        toast.error("Missing file data for download");
        return;
      }

      try {
        toast.loading("Preparing download...");
        let base64Data = file.documentbody;

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
        const blob = new Blob([byteArray], { type: file.mimetype });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", file.filename);
        document.body.appendChild(link);
        link.click();
        if (link.parentNode) {
          link.parentNode.removeChild(link);
        }
        URL.revokeObjectURL(url);
        toast.dismiss();
        toast.success(`${file.filename} downloaded successfully!`);
      } catch (error) {
        toast.error(`Failed to download file: ${(error as Error).message}`);
      }
    } else {
      if ("absoluteurl" in file && file.absoluteurl) {
        try {
          const link = document.createElement("a");
          link.href = file.absoluteurl;
          link.target = "_blank";
          link.rel = "noopener noreferrer";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);

          toast.success(`Opening ${file.fullname} in a new tab.`);
        } catch (error) {
          console.error("Error opening the file:", error);
          toast.error(
            `Failed to open file: ${
              error instanceof Error ? error.message : "Unknown error"
            }`
          );
        }
      } else {
        toast.error("Missing URL for SharePoint file download");
      }
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

  removeFile = async (fileId?: string) => {
    if (!fileId) {
      console.error("No file ID provided for deletion");
      toast.error("No file ID provided");
      return;
    }

    if (!this.state.sharePointDocLoc) {
      this.setState((prevState) => ({
        files: prevState.files.map((file) =>
          file.noteId === fileId ? { ...file, isLoading: true } : file
        ),
      }));

      toast.promise(deleteRelatedNote(this.props.context!, fileId), {
        loading: "Deleting file...",
        success: () => {
          this.setState((prevState) => ({
            files: prevState.files.filter((file) => file.noteId !== fileId),
          }));
          return "File successfully deleted!";
        },
        error: (err) => {
          this.setState((prevState) => ({
            files: prevState.files.map((file) =>
              file.noteId === fileId ? { ...file, isLoading: false } : file
            ),
          }));
          console.error("Toast Error: ", err);
          return `Failed to delete file: ${err.message || "Unknown error"}`;
        },
      });
    } else {
      const file = this.state.sharePointData.find(
        (f: SharePointDocument) => f.sharepointdocumentid === fileId
      );
      if (!file) {
        toast.error("SharePoint document not found.");
        return;
      }

      toast.promise(
        deleteSharePointDocument(
          this.props.context!,
          file.sharepointdocumentid,
          file.documentid,
          file.filetype,
          this.state.selectedDocumentLocation!
        ),
        {
          loading: "Deleting SharePoint document...",
          success: () => {
            this.setState((prevState) => ({
              sharePointData: prevState.sharePointData.filter(
                (f) => f.sharepointdocumentid !== fileId
              ),
            }));
            return "SharePoint document successfully deleted!";
          },
          error: (err) => {
            console.error("Error deleting SharePoint document: ", err);
            return `Failed to delete SharePoint document: ${
              err.message || "Unknown error"
            }`;
          },
        }
      );
    }
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
  createSharePointFolder = () => {
    const { context } = this.props;
    const { newFolderName, currentFolderPath } = this.state;
    if (!newFolderName) {
      toast.error("Folder name cannot be empty.");
      return;
    }

    const folderCreationPromise = createSharePointFolder(
      context!,
      newFolderName,
      currentFolderPath,
      this.state.selectedDocumentLocation!
    );

    toast.promise(folderCreationPromise, {
      loading: "Creating folder...",
      success: (data) => {
        this.toggleModal();
        this.loadExistingFiles();
        this.setState({ newFolderName: "" });
        return `Folder '${newFolderName}' created successfully!`;
      },
      error: (err) => {
        return `Failed to create folder: ${err.message || err.toString()}`;
      },
    });
  };

  performActionOnSelectedFiles = (action: FileAction) => {
    this.state.selectedFiles.forEach((fileId) => {
      let file;

      if (this.state.sharePointDocLoc) {
        file = this.state.sharePointData.find(
          (f: SharePointDocument) => f.sharepointdocumentid === fileId
        );
      } else {
        file = this.state.files.find((f: FileData) => f.noteId === fileId);
      }

      if (!file) {
        toast.error("File not found.");
        return;
      }

      if (this.state.sharePointDocLoc && "sharepointdocumentid" in file) {
        // Actions for SharePoint documents
        switch (action) {
          case "download":
            this.downloadFile(file);
            break;
          case "delete":
            this.removeFile(file.sharepointdocumentid);
            break;
          default:
            toast.error(
              `Action ${action} is not supported for SharePoint documents.`
            );
        }
      } else if (!this.state.sharePointDocLoc && "noteId" in file) {
        // Actions for notes
        switch (action) {
          case "duplicate":
            this.duplicateFile(file.noteId!);
            break;
          case "download":
            this.downloadFile(file);
            break;
          case "delete":
            this.removeFile(file.noteId);
            break;
          case "edit":
            this.toggleEditModal(file.noteId);
            break;
          case "preview":
            this.openDialog({
              filename: file.filename,
              documentbody: file.documentbody!,
              mimetype: file.mimetype!,
            });
            break;
          default:
            toast.error(`Action ${action} is not supported for notes.`);
        }
      }
    });

    this.setState({ selectedFiles: [] });
  };

  openDialog(file: PreviewFile) {
    this.setState({
      isDialogOpen: true,
      previewFile: {
        ...file,
        documentbody: createDataUri(file.mimetype, file.documentbody)
      },
      xlsxContent: '',
      xlsxData: null
    });
  }

  closeDialog() {
    this.setState({ isDialogOpen: false, previewFile: null });
  }

  handleSearch = (event: any, newValue?: string) => {
    const searchText = newValue || "";

    if (this.state.sharePointDocLoc) {
      this.setState({ spSearchText: searchText });
    } else {
      this.setState({ notesSearchText: searchText });
    }
  };

  toggleSortOrder = () => {
    this.setState((prevState) => ({ sortAsc: !prevState.sortAsc }));
  };
  toggleTooltip() {
    this.setState((prevState) => ({ showTooltip: !prevState.showTooltip }));
  }

  toggleSharePointDocLoc(_: React.MouseEvent<HTMLElement>, checked?: boolean) {
    this.setState({ sharePointDocLoc: !!checked, selectedFiles: [] }, () => {
      this.loadExistingFiles();
    });
  }
  getFilteredAndSortedFiles() {
    const { files, notesSearchText } = this.state;
    return files.filter((file) =>
      file.filename.toLowerCase().includes(notesSearchText.toLowerCase())
    );
  }

  getFilteredAndSortedSPFiles() {
    const { sharePointData, spSearchText } = this.state;
    return sharePointData
      .filter((spData) =>
        spData.fullname
          .toLowerCase()
          .includes(spSearchText.toLowerCase().trim())
      )
      .sort((a, b) => {
        if (a.filetype === "folder" && b.filetype !== "folder") {
          return -1;
        } else if (a.filetype !== "folder" && b.filetype === "folder") {
          return 1;
        }
        return a.fullname.toLowerCase().localeCompare(b.fullname.toLowerCase());
      });
  }

  handleDropdownChange(
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (option) {
      this.setState({
        selectedDocumentLocation: option.key as string,
        selectedDocumentLocationName: option.text as string,
      });
      this.setState({ currentFolderPath: "" }, () => {
        this.loadExistingFiles();
      });
    }
  }

  renderRibbon = () => {
    const { selectedFiles, sharePointDocLoc, isCollapsed } = this.state;
    const commandButtons = (
      <Stack horizontal tokens={ribbonStackTokens}>
        {!sharePointDocLoc && selectedFiles.length < 2 && (
          <>
            <CommandButton
              text="Preview"
              iconProps={{ iconName: "View" }}
              onClick={() => this.performActionOnSelectedFiles("preview")}
              disabled={
                selectedFiles.length !== 1 || 
                !this.state.files.some(file => 
                  selectedFiles.includes(file.noteId!) && 
                  (isImage(file.mimetype!) || isPDF(file.mimetype!))
                )
              }
              className="icon-button"
            />
            <CommandButton
              iconProps={{ iconName: "Edit" }}
              text="Rename"
              onClick={() => this.performActionOnSelectedFiles("edit")}
              disabled={selectedFiles.length === 0}
              className="icon-button"
            />
          </>
        )}
        <CommandButton
          iconProps={{ iconName: "Download" }}
          text="Download"
          onClick={() => this.performActionOnSelectedFiles("download")}
          disabled={selectedFiles.length === 0}
          className="icon-button"
        />
        {!sharePointDocLoc && (
          <CommandButton
            iconProps={{ iconName: "Copy" }}
            text="Duplicate"
            onClick={() => this.performActionOnSelectedFiles("duplicate")}
            disabled={selectedFiles.length === 0}
            className="icon-button"
          />
        )}
        <CommandButton
          iconProps={{ iconName: "Delete" }}
          text="Delete"
          onClick={() => this.performActionOnSelectedFiles("delete")}
          disabled={selectedFiles.length === 0}
          className="icon-button"
        />
      </Stack>
    );
    const menuItems = [
      {
        key: "preview",
        text: "Preview",
        iconProps: { iconName: "View" },
        onClick: () => this.performActionOnSelectedFiles("preview"),
        disabled:
          selectedFiles.length !== 1 || 
          !this.state.files.some(file => 
            selectedFiles.includes(file.noteId!) && 
            (isImage(file.mimetype!) || isPDF(file.mimetype!) || isExcel(file.mimetype!))
          )
      },
      {
        key: "rename",
        text: "Rename",
        iconProps: { iconName: "Edit" },
        onClick: () => this.performActionOnSelectedFiles("edit"),
        disabled: selectedFiles.length === 0 || selectedFiles.length >= 2,
      },
      {
        key: "download",
        text: "Download",
        iconProps: { iconName: "Download" },
        onClick: () => this.performActionOnSelectedFiles("download"),
        disabled: selectedFiles.length === 0,
      },
      {
        key: "duplicate",
        text: "Duplicate",
        iconProps: { iconName: "Copy" },
        onClick: () => this.performActionOnSelectedFiles("duplicate"),
        disabled: selectedFiles.length === 0,
      },
      {
        key: "delete",
        text: "Delete",
        iconProps: { iconName: "Delete" },
        onClick: () => this.performActionOnSelectedFiles("delete"),
        disabled: selectedFiles.length === 0,
      },
    ];
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
                <>
                  {this.state.menuVisible && (
                    <ContextualMenu
                      items={menuItems}
                      target={this.state.target}
                      onDismiss={this.closeMenu}
                      isBeakVisible={true}
                    />
                  )}
                  {isCollapsed ? (
                    <>
                      <CommandButton
                        text="Actions"
                        onClick={this.toggleMenu}
                        className="icon-button"
                      />
                    </>
                  ) : (
                    commandButtons
                  )}
                </>
              </Stack>
            )}
            {selectedFiles.length < 1 && sharePointDocLoc === true && (
              <CommandButton
                iconProps={{ iconName: "folder" }}
                text="Create Folder"
                onClick={() => this.toggleModal()}
                disabled={selectedFiles.length > 0}
                className="icon-button"
              />
            )}
          </Stack>
        </Stack>
      </>
    );
  };

  toggleFileSelection = (
    fileId: string,
    file?: SharePointDocument,
    forceSelect: boolean = false
  ) => {
    if (file && file.filetype === "folder") {
      this.handleFolderClick(file.relativelocation);
      return;
    }

    const isSelected = this.state.selectedFiles.includes(fileId);

    if (forceSelect && !isSelected) {
      this.setState((prevState) => ({
        selectedFiles: [...prevState.selectedFiles, fileId],
      }));
    } else if (!forceSelect) {
      this.setState((prevState) => ({
        selectedFiles: isSelected
          ? prevState.selectedFiles.filter((id) => id !== fileId)
          : [...prevState.selectedFiles, fileId],
      }));
    }
  };

  toggleModal = () => {
    this.setState((prevState) => ({ showModal: !prevState.showModal }));
  };

  handleFolderNameChange = (
    _: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    this.setState({ newFolderName: newValue || "" });
  };


  renderFileList() {
    const { sharePointDocLoc, selectedFiles, currentFolderPath } = this.state;
    const files = this.getFilteredAndSortedFiles();
    const sharePointData = this.getFilteredAndSortedSPFiles();
    if (sharePointDocLoc) {
      return (
        <div style={{ display: "flex", alignItems: "center" }}>
          {currentFolderPath && (
            <div style={{ marginRight: "20px" }}>
              <IconButton
                iconProps={{ iconName: "Back" }}
                title="Back"
                ariaLabel="Go back"
                onClick={(event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  this.handleBackClick();
                }}
              />
            </div>
          )}
          <div style={{ display: "flex", flexWrap: "wrap", flexGrow: 1 }}>
            {sharePointData.map((file) => (
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
                onContextMenu={(event) => {
                  if (event.button === 2 && file.sharepointdocumentid) {
                    event.preventDefault();
                    const fileId = file.sharepointdocumentid;
                    const isSelected =
                      this.state.selectedFiles.includes(fileId);
                    this.toggleFileSelection(
                      file.sharepointdocumentid,
                      file,
                      isSelected
                    );
                    this.toggleMenu(event);
                  }
                }}
              >
                <div className="file-image">
                  {file.filetype === "folder" ? (
                    <FolderIcon />
                  ) : (
                    <Icon
                      {...getFileTypeIconProps({
                        extension: this.getFileExtension(file.fullname),
                        size: 96,
                        imageFileType: "svg",
                      })}
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
            ))}
          </div>
        </div>
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
          onContextMenu={(event) => {
            if (event.button === 2 && file.noteId) {
              event.preventDefault();
              const fileId = file.noteId;
              const isSelected = this.state.selectedFiles.includes(fileId);
              this.toggleFileSelection(fileId!, undefined, isSelected);
              this.toggleMenu(event);
            }
          }}
        >
          <Icon
            {...getFileTypeIconProps({
              extension: this.getFileExtension(file.filename),
              size: 96,
              imageFileType: "svg",
            })}
          />
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
      showModal,
      newFolderName,
      isDialogOpen,
      previewFile,
      xlsxContent
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
        closeDelay={500}
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

        {isDialogOpen && previewFile && (
          <Modal
            isOpen={isDialogOpen}
            onDismiss={this.closeDialog}
            isBlocking={false}
            containerClassName="ms-modalExample-container"
            styles={{
              main: {
                width: "100%",
                height: "100%",
                "@media(max-width: 767px)": { width: "100%" },
              },
              scrollableContent: {
                height: "100%",
                overflow: "hidden",
              },
            }}
          >
            <div className={"scrollable-area"}>
              {isPDF(previewFile.mimetype) && (
                <object
                  data={previewFile.documentbody.replace(/(data:application\/pdf;base64,).*?\1/, '$1')}
                  type="application/pdf"
                  width="100%"
                  height="100%"
                >
                  <p>
                    Your browser does not support PDFs.{" "}
                    <a href={previewFile.documentbody}>Download the PDF</a>.
                  </p>
                </object>
              )}
              {isExcel(previewFile.mimetype) && (
                  <div id="xlsx-preview" style={{ height: '100%' }}></div>
                )}
              {isImage(previewFile.mimetype) && (
                <Img
                  src={previewFile.documentbody.replace(/(data:image\/[a-zA-Z]+;base64,).*?\1/, '$1')}
                  alt={previewFile.filename}
                />
              )}
            </div>

            <DialogFooter
              styles={{
                actionsRight: {
                  marginTop: "-6px",
                  marginRight: "9px",
                  marginBottom: "0px",
                },
              }}
            >
              <DefaultButton onClick={this.closeDialog} text="Close" />
            </DialogFooter>
          </Modal>
        )}

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
        <Dialog
          hidden={!showModal}
          onDismiss={this.toggleModal}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Create New Folder",
            subText: "Enter a name of the new folder:",
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 450 } },
          }}
        >
          <TextField
            value={newFolderName}
            onChange={this.handleFolderNameChange}
            placeholder="Folder name"
          />
          <DialogFooter>
            <PrimaryButton
              onClick={this.createSharePointFolder}
              disabled={!newFolderName.trim()}
              text="Create"
            />
            <DefaultButton onClick={this.toggleModal} text="Cancel" />
          </DialogFooter>
        </Dialog>
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
          <Stack
            horizontal
            style={{
              width: "100%",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
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
          </Stack>
        </div>
      </>
    );
  }
}

export default Landing;
