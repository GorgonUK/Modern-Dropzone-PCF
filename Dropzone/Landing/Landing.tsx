import React, { Component } from "react";
import Dropzone from "react-dropzone";
import { IInputs } from "../generated/ManifestTypes";
import "../css/Dropzone.css";
import {
  createRelatedNote,
  getRelatedNotes,
  updateRelatedNote,
  deleteRelatedNote,
  getSharePointLocations,
  getSharePointFolderData,
  createSharePointDocument,
  deleteSharePointDocument,
  createSharePointFolder,
  createSharePointLocation,
  createAtivityDocument,
} from "../DataverseActions";
import {
  FileData,
  SharePointDocument,
  PreviewFile,
  GenericActionResponse,
} from "../Interfaces";
import { Img } from "react-image";
import {
  isPDF,
  isImage,
  isExcel,
  createDataUri,
  isActivityType,
  focusSPDocumentsAndRestore,
} from "../utils";
import {
  IContextualMenuItem,
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
  IIconProps,
  Callout,
  Label,
  HighContrastSelector,
  DirectionalHint,
  ComboBox,
  IComboBox,
  IComboBoxOption,
  IButtonProps,
} from "@fluentui/react";
import {
  getFileTypeIconProps,
  initializeFileTypeIcons,
} from "@uifabric/file-type-icons";
import { read, utils } from "xlsx";
import { Tooltip } from "react-tippy";
import "react-tippy/dist/tippy.css";
import toast, { Toaster } from "react-hot-toast";
import FolderIcon from "./FolderIcon";
import { getEntityMetadata, getControlValue } from "../utils";

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
  documentLocations: {
    name: string;
    sharepointdocumentlocationid: string;
    parentsiteid?: string;
  }[];
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
  sharePointEnabled: boolean;
  userPreference: boolean | undefined;
  isCalloutVisible: boolean;
  isFolderDeletionDialogVisible: boolean;
  isCreateLocationDialogVisible: boolean;
  createLocationDisplayName: string;
  selectedCreateLocation: string;
  createLocationFolderName: string;
  isSaveButtonEnabled: boolean;
  isActivity: boolean;
  sharePointEnabledParameter: boolean;
  selectedFolderForDelete: SharePointDocument | null;
  formType: number;
}

type FileAction =
  | "edit"
  | "download"
  | "delete"
  | "preview"
  | "addToActivityAttachment";

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
  const buffer = await fetch(documentbody).then((res) => res.arrayBuffer());
  const workbook = read(buffer, { type: "array" });
  const sheetsData: SheetData = {};

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    sheetsData[sheetName] = utils.sheet_to_json(worksheet, { header: 1 });
  });

  return sheetsData;
}

export class Landing extends Component<LandingProps, LandingState> {
  private targetRef: React.RefObject<HTMLDivElement>;

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
      xlsxContent: "",
      xlsxData: null,
      sharePointEnabled: false,
      userPreference: false,
      isCalloutVisible: false,
      isFolderDeletionDialogVisible: false,
      isCreateLocationDialogVisible: false,
      createLocationDisplayName: "",
      selectedCreateLocation: "",
      createLocationFolderName: "",
      isSaveButtonEnabled: false,
      isActivity: false,
      sharePointEnabledParameter: false,
      selectedFolderForDelete: null,
      formType: 0,
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
    this.onGearIconClick = this.onGearIconClick.bind(this);
    this.onCalloutDismiss = this.onCalloutDismiss.bind(this);
    this.toggleRememberLocation = this.toggleRememberLocation.bind(this);
    this.handleRemoveFolderClick = this.handleRemoveFolderClick.bind(this);
    this.isActivityType = this.isActivityType.bind(this);
    this.addFileAttachmentToActivity =
      this.addFileAttachmentToActivity.bind(this);
    this.targetRef = React.createRef();
  }

  handleRemoveFolderClick = (event: any, file: SharePointDocument): void => {
    //console.log(file.title);
    event.preventDefault();
    event.stopPropagation();

    this.setState({ selectedFolderForDelete: file }, () => {
      this.toggleRemoveFolderDialog();
    });
  };

  toggleRemoveFolderDialog = (): void => {
    this.setState((prevState) => ({
      isFolderDeletionDialogVisible: !prevState.isFolderDeletionDialogVisible,
    }));
  };

  private handlecreateLocation = async () => {
    const {
      createLocationDisplayName,
      createLocationFolderName,
      selectedCreateLocation,
    } = this.state;
    const { context } = this.props;

    const createLocationPromise = createSharePointLocation(
      context!,
      createLocationDisplayName,
      createLocationFolderName,
      selectedCreateLocation
    );

    toast.promise(createLocationPromise, {
      loading: "Creating location...",
      success: (res) => {
        this.setState({ isCreateLocationDialogVisible: false });
        const createdSPLocation: IDropdownOption = {
          key: res,
          text: createLocationDisplayName,
        };
        this.handleDropdownChange(
          {} as React.FormEvent<HTMLDivElement>,
          createdSPLocation
        );

        this.setState((prevState) => ({
          documentLocations: [
            ...prevState.documentLocations,
            {
              name: createLocationDisplayName,
              sharepointdocumentlocationid: res,
            },
          ],
          selectedDocumentLocation: res,
        }));

        return `Location "${createLocationDisplayName}" created successfully!`;
      },
      error: "Error creating location",
    });
  };

  private handleInputChange = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ) => {
    const target = event.currentTarget;
    const name = target.name as keyof LandingState;
    const value = newValue ?? target.value;
    this.setState(
      (prevState) =>
        ({
          ...prevState,
          [name]: value,
        } as unknown as LandingState),
      this.validateForm
    );
  };
  private validateForm = (): void => {
    const {
      createLocationDisplayName,
      selectedCreateLocation,
      createLocationFolderName,
    } = this.state;
    const isSaveButtonEnabled =
      createLocationDisplayName.trim() !== "" &&
      selectedCreateLocation.trim() !== "" &&
      createLocationFolderName.trim() !== "";
    this.setState({ isSaveButtonEnabled });
  };
  private handleCreateLocationDropdownChange = (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption
  ): void => {
    if (option) {
      //console.log(option.key);
      this.setState(
        { selectedCreateLocation: option.key as string },
        this.validateForm
      );
    }
  };

  private openCreateLocationDialog = (): void => {
    this.setState({ isCreateLocationDialogVisible: true });
  };

  private closeCreateLocationDialog = (): void => {
    this.setState({ isCreateLocationDialogVisible: false });
  };
  private targetElement: HTMLElement | null = null;
  private onGearIconClick = (
    ev?:
      | React.MouseEvent<HTMLElement, MouseEvent>
      | React.KeyboardEvent<HTMLElement>,
    item?: IContextualMenuItem
  ): boolean | void => {
    if (ev && ev.currentTarget instanceof HTMLElement) {
      this.targetElement = ev.currentTarget;
      this.toggleCallout();
    }
  };

  private toggleCallout = (): void => {
    this.setState((prevState) => ({
      isCalloutVisible: !prevState.isCalloutVisible,
      target: this.targetElement,
    }));
  };
  private onCalloutDismiss = (): void => {
    this.setState({ isCalloutVisible: false });
  };
  private async handleFolderClick(folderPath: string): Promise<void> {
    this.setState(
      (prevState) => ({
        folderStack: [...prevState.folderStack, prevState.currentFolderPath],
        isLoading: true,
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
        } finally {
          this.setState({ isLoading: false });
        }
      }
    );
  }

  private handleBackClick = async () => {
    this.setState(
      (prevState) => {
        const folderStack = [...prevState.folderStack];
        const previousFolderPath = folderStack.pop();
        return {
          folderStack,
          currentFolderPath: previousFolderPath || "",
          isLoading: true,
        };
      },
      async () => {
        try {
          const { currentFolderPath } = this.state;
          const data = await getSharePointFolderData(
            this.props.context!,
            currentFolderPath,
            this.state.selectedDocumentLocation!,
            this.state.selectedDocumentLocationName
          );
          this.setState({ sharePointData: data });
        } catch (error) {
          console.error("Error fetching SharePoint folder data:", error);
        } finally {
          this.setState({ isLoading: false });
        }
      }
    );
  };

  private async getSharePointLocations(): Promise<void> {
    const { context } = this.props;
    try {
      const response = await getSharePointLocations(context!);
      const firstLocationId =
        response.length > 0 ? response[0].sharepointdocumentlocationid : "";
      this.setState((prevState) => ({
        documentLocations: response,
        selectedDocumentLocation:
          prevState.selectedDocumentLocation || firstLocationId,
        selectedDocumentLocationName:
          prevState.selectedDocumentLocationName ||
          (response.length > 0 ? response[0].name : ""),
      }));
    } catch (error) {
      console.error("Error fetching SharePoint document locations:", error);
    }
  }

  private isDefaultLocation() {
    if (
      this.state.documentLocations.length === 1 &&
      this.state.selectedDocumentLocationName === "Documents on Default Site 1"
    ) {
      return false;
    } else {
      return true;
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
        let defaultLocation = this.isDefaultLocation();
        const data = await getSharePointFolderData(
          this.props.context!,
          currentFolderPath,
          this.state.selectedDocumentLocation!,
          this.state.selectedDocumentLocationName,
          defaultLocation
        );
        this.setState({ sharePointData: data });
      }
    );
  }

  async isActivityType(): Promise<boolean> {
    const metadata = await getEntityMetadata(this.props.context!);
    if (!metadata) {
      return false;
    }
    const isActivity = isActivityType(metadata.schemaName);
    if (isActivity) {
      return true;
    } else {
      return false;
    }
  }

  async getUserSettings(): Promise<any> {
    const settings = localStorage.getItem("userSettings");
    return settings ? JSON.parse(settings) : null;
  }

  async checkUserSettings(): Promise<boolean | undefined> {
    const settings = localStorage.getItem("userSettings");

    if (settings) {
      const parsedSettings = JSON.parse(settings);
      const metadata = await getEntityMetadata(this.props.context!);
      if (
        parsedSettings.tableName !== metadata?.schemaName ||
        parsedSettings.entityId !== metadata?.entityId
      ) {
        return undefined;
      }
      if (parsedSettings.tableName === metadata?.schemaName) {
        if (typeof parsedSettings.NotesOrSharePoint === "undefined") {
          return parsedSettings.NotesOrSharePoint;
        } else {
          return undefined;
        }
      } else {
        return false;
      }
    } else {
      return false;
    }
  }

  async checkSharePointIntegration(): Promise<boolean> {
    const response = await getSharePointLocations(this.props.context!);
    return response.length === 0 ? true : false;
  }

  async componentDidMount() {
    var formType = Xrm.Page.ui.getFormType();
    this.setState({ formType });

    // Onsave handler, works well with Active and Deactive ribbon buttons
    Xrm.Page.data.entity.addOnSave(() => {
      const interval = 500;
      const duration = 5000;
      const endTime = Date.now() + duration;

      const intervalId = setInterval(() => {
        const formType = Xrm.Page.ui.getFormType();
        this.setState({ formType });

        if (Date.now() >= endTime) {
          clearInterval(intervalId);
        }
      }, interval);
    });

    window.addEventListener("recordSavedEvent", async (event: any) => {
      //console.log("save event");
      const passedEntityId = event.detail.entityId;
      const checkEntityId = async (): Promise<void> => {
        const entityId = (this.props.context as any).page.entityId;
        if (entityId && entityId === passedEntityId) {
          if (this.state.sharePointEnabledParameter == true) {
            focusSPDocumentsAndRestore();
          }
        } else {
          await this.delay(2000);
          return checkEntityId();
        }
      };
      await checkEntityId();
    });
    const sharePointEnabled = await this.checkSharePointIntegration();
    const userPreference = await this.checkUserSettings();
    const settings = await this.getUserSettings();
    const isActivity = await this.isActivityType();
    let sharePointEnabledParameter =
      getControlValue(this.props.context!, "enableSharePointDocuments") ===
      true;
    this.setState({ sharePointEnabledParameter });

    if (
      settings &&
      settings.selectedDocumentLocation &&
      settings.selectedDocumentLocationName &&
      (sharePointEnabledParameter === true || sharePointEnabled == false)
    ) {
      this.setState(
        {
          selectedDocumentLocation: settings.selectedDocumentLocation,
          selectedDocumentLocationName: settings.selectedDocumentLocationName,
        },
        () => {
          this.setState(
            {
              sharePointEnabled: sharePointEnabled,
              userPreference,
              isActivity,
              sharePointDocLoc: true,
            },
            () => {
              this.loadExistingFiles().then(() => {
                this.setState({ isLoading: false });
              });
            }
          );
        }
      );
    } else {
      this.setState(
        {
          sharePointEnabled: sharePointEnabled,
          userPreference,
          isActivity,
        },
        () => {
          this.loadExistingFiles().then(() => {
            this.setState({ isLoading: false });
          });
        }
      );
    }

    window.addEventListener("resize", this.handleResize);
  }

  componentWillUnmount() {
    window.removeEventListener("resize", this.handleResize);
  }

  handleResize() {
    const element = document.querySelector(".ribbon-dropzone-wrapper");
    if (element) {
      const wrapper = element.getBoundingClientRect();
      this.setState({ isCollapsed: wrapper.width < 703 });
    }
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
    if (this.state.formType !== 2) {
      return;
    }
    let allowNoteDropsParameter =
      getControlValue(this.props.context!, "allowNoteDrops") === true;

    let allowSharePointDropsParameter =
      getControlValue(this.props.context!, "allowSharePointDrops") === true;

    if (!this.state.sharePointDocLoc && allowNoteDropsParameter) {
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
    } else if (allowSharePointDropsParameter) {
      acceptedFiles.forEach((file) => {
        const reader = new FileReader();
        reader.onload = () => {
          const binaryStr = reader.result as string;
          let defaultLocation;
          if (
            this.state.documentLocations.length == 1 &&
            this.state.selectedDocumentLocationName ==
              "Documents on Default Site 1"
          ) {
            defaultLocation = "";
          } else {
            defaultLocation = this.state.selectedDocumentLocation!;
          }
          const uploadPromise = createSharePointDocument(
            this.props.context!,
            file.name,
            binaryStr,
            this.state.currentFolderPath,
            defaultLocation
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

    try {
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
          await this.getSharePointLocations();
          await this.getSharePointData();
        } else {
          await this.getSharePointData();
        }
      }
    } catch (error) {
      console.error("Error loading files:", error);
    } finally {
      await this.delay(2000);
      this.setState({ isLoading: false });
    }
  };

  delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  addFileAttachmentToActivity = async (file: FileData | SharePointDocument) => {
    //console.log(file);
    if (
      !("documentbody" in file) ||
      !file.documentbody ||
      !file.mimetype ||
      !file.filename
    ) {
      toast.error("Missing file data for attachment");
      return;
    }

    try {
      toast.loading("Attaching file...");

      const response: GenericActionResponse = await createAtivityDocument(
        this.props.context!,
        file.filename,
        file.documentbody,
        file.mimetype
      );

      toast.dismiss();
      if (response.success) {
        toast.success(`${file.filename} attached successfully!`);
      } else {
        toast.error(response.message as string);
      }
    } catch (error) {
      toast.dismiss();
      toast.error(`Failed to attach file: ${(error as Error).message}`);
    }
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
      this.setState({ isLoading: true });

      const file = this.state.sharePointData.find(
        (f: SharePointDocument) => f.sharepointdocumentid === fileId
      );
      if (!file) {
        this.setState({ isLoading: false });
        toast.error("SharePoint document not found.");
        return;
      }
      let defaultLocation = this.isDefaultLocation();
      toast.promise(
        deleteSharePointDocument(
          this.props.context!,
          file.sharepointdocumentid,
          file.documentid,
          file.filetype,
          this.state.selectedDocumentLocation!,
          defaultLocation
        ),
        {
          loading: "Deleting SharePoint document...",
          success: () => {
            this.setState((prevState) => ({
              sharePointData: prevState.sharePointData.filter(
                (f) => f.sharepointdocumentid !== fileId
              ),
              isLoading: false,
            }));
            return "SharePoint document successfully deleted!";
          },
          error: (err) => {
            this.setState({ isLoading: false });
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

  middleEllipsis(
    filename: string,
    maxLength: number = 18,
    isFolder: boolean = false
  ): string {
    if (filename.length < maxLength) {
      return filename;
    }

    if (isFolder) {
      const startChars = maxLength - 3;
      const start = filename.substring(0, startChars);

      return `${start}...`;
    } else {
      const lastDotIndex = filename.lastIndexOf(".");
      if (lastDotIndex === -1) {
        const startChars = maxLength - 3;
        const start = filename.substring(0, startChars);

        return `${start}...`;
      }

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
  }

  createSharePointFolder = () => {
    const { context } = this.props;
    const { newFolderName, currentFolderPath } = this.state;
    if (!newFolderName) {
      toast.error("Folder name cannot be empty.");
      return;
    }
    let defaultLocation = this.isDefaultLocation();
    const folderCreationPromise = createSharePointFolder(
      context!,
      newFolderName,
      currentFolderPath,
      this.state.selectedDocumentLocation!,
      defaultLocation
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
          case "addToActivityAttachment":
            this.addFileAttachmentToActivity(file);
            break;
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
        documentbody: createDataUri(file.mimetype, file.documentbody),
      },
      xlsxContent: "",
      xlsxData: null,
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
  async saveUserSettings(
    sharePointEnabled: boolean,
    selectedDocumentLocation: string | null,
    selectedDocumentLocationName: string
  ) {
    const metadata = await getEntityMetadata(this.props.context!);
    const settings = {
      tableName: metadata?.schemaName,
      NotesOrSharePoint: sharePointEnabled,
      selectedDocumentLocation: selectedDocumentLocation,
      selectedDocumentLocationName: selectedDocumentLocationName,
    };
    localStorage.setItem("userSettings", JSON.stringify(settings));
  }
  toggleSharePointDocLoc = async (
    _: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ) => {
    const sharePointDocLoc = !!checked;
    const { selectedDocumentLocation, selectedDocumentLocationName } =
      this.state;
    this.setState(
      { sharePointDocLoc, selectedFiles: [], isLoading: true },
      async () => {
        this.loadExistingFiles();
      }
    );
  };
  toggleRememberLocation = async (
    _: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ) => {
    const sharePointDocLoc = !!checked;
    const { selectedDocumentLocation, selectedDocumentLocationName } =
      this.state;
    this.setState({ userPreference: !!checked }, async () => {
      await this.saveUserSettings(
        sharePointDocLoc,
        selectedDocumentLocation,
        selectedDocumentLocationName
      );
    });
  };
  getFilteredAndSortedFiles() {
    const { files, notesSearchText } = this.state;
    return files.filter(
      (file) =>
        file.filename &&
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

  async handleDropdownChange(
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> {
    const userPreference = false;
    if (option) {
      const selectedDocumentLocation = option.key as string;
      const selectedDocumentLocationName = option.text as string;

      this.setState(
        {
          userPreference,
          selectedDocumentLocation,
          selectedDocumentLocationName,
        },
        async () => {
          await this.saveUserSettings(
            userPreference,
            selectedDocumentLocation,
            selectedDocumentLocationName
          );
          this.setState({ currentFolderPath: "" }, () => {
            this.loadExistingFiles();
          });
        }
      );
    }
  }

  renderRibbon = () => {
    const { selectedFiles, sharePointDocLoc, isCollapsed, formType } =
      this.state;
    const formIsDisabled = formType !== 2;
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
                !this.state.files.some(
                  (file) =>
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
        <CommandButton
          iconProps={{ iconName: "Delete" }}
          text="Delete"
          onClick={() => this.performActionOnSelectedFiles("delete")}
          disabled={selectedFiles.length === 0 || formIsDisabled}
          className="icon-button"
        />
      </Stack>
    );

    const menuItems: IContextualMenuItem[] = [
      {
        key: "download",
        text: "Download",
        iconProps: { iconName: "Download" },
        onClick: () => this.performActionOnSelectedFiles("download"),
        disabled: selectedFiles.length === 0,
      },
      {
        key: "delete",
        text: "Delete",
        iconProps: { iconName: "Delete" },
        onClick: () => this.performActionOnSelectedFiles("delete"),
        disabled: selectedFiles.length === 0 || formIsDisabled,
      },
    ];

    if (!sharePointDocLoc) {
      menuItems.unshift(
        {
          key: "preview",
          text: "Preview",
          iconProps: { iconName: "View" },
          onClick: () => this.performActionOnSelectedFiles("preview"),
          disabled:
            selectedFiles.length !== 1 ||
            !this.state.files.some(
              (file) =>
                selectedFiles.includes(file.noteId!) &&
                (isImage(file.mimetype!) || isPDF(file.mimetype!))
            ),
        },
        {
          key: "rename",
          text: "Rename",
          iconProps: { iconName: "Edit" },
          onClick: () => this.performActionOnSelectedFiles("edit"),
          disabled:
            selectedFiles.length === 0 ||
            selectedFiles.length >= 2 ||
            formIsDisabled,
        }
      );
    }
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
            {selectedFiles.length < 1 &&
              sharePointDocLoc === true &&
              !formIsDisabled && (
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
    const {
      sharePointDocLoc,
      selectedFiles,
      currentFolderPath,
      isFolderDeletionDialogVisible,
      selectedFolderForDelete,
      formType,
    } = this.state;
    const files = this.getFilteredAndSortedFiles();
    const sharePointData = this.getFilteredAndSortedSPFiles();
    const formIsDisabled = formType !== 2;

    if (sharePointDocLoc) {
      return (
        <div style={{ display: "flex", alignItems: "center" }}>
          {currentFolderPath && sharePointData.length < 0 && (
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
                  event.preventDefault();
                  if (
                    file.filetype !== "folder" &&
                    event.button === 2 &&
                    file.sharepointdocumentid
                  ) {
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
                <div className="file-image" style={{ position: "relative" }}>
                  {file.filetype === "folder" ? (
                    <>
                      <FolderIcon />
                      {!formIsDisabled && (
                        <IconButton
                          className="remove-button"
                          iconProps={{ iconName: "Cancel" }}
                          title="Remove"
                          ariaLabel="Remove"
                          onClick={(event) =>
                            this.handleRemoveFolderClick(event, file)
                          }
                        />
                      )}
                      <Dialog
                        hidden={!isFolderDeletionDialogVisible}
                        onDismiss={this.toggleRemoveFolderDialog}
                        dialogContentProps={{
                          type: DialogType.normal,
                          title: "Remove Folder",
                          subText: `Are you sure you want to delete folder "${selectedFolderForDelete?.fullname}"?`,
                        }}
                      >
                        <DialogFooter>
                          <PrimaryButton
                            onClick={() => {
                              this.removeFile(
                                selectedFolderForDelete?.sharepointdocumentid
                              );
                              this.toggleRemoveFolderDialog();
                            }}
                            text="Yes"
                          />
                          <DefaultButton
                            onClick={this.toggleRemoveFolderDialog}
                            text="No"
                          />
                        </DialogFooter>
                      </Dialog>
                    </>
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
                    {this.middleEllipsis(file.fullname, 18, true)}
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
      return (
        <div style={{ display: "flex", flexWrap: "wrap", flexGrow: 1 }}>
          {files.map((file) => (
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
                <p className="file-name">
                  {this.middleEllipsis(file.filename)}
                </p>
              </Tooltip>
              <p className="file-size">{this.formatFileSize(file.filesize)}</p>
            </div>
          ))}
        </div>
      );
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
      sharePointData,
      sharePointEnabled,
      currentFolderPath,
      isCalloutVisible,
      isCreateLocationDialogVisible,
      userPreference,
      createLocationDisplayName,
      selectedCreateLocation,
      createLocationFolderName,
      isSaveButtonEnabled,
      sharePointEnabledParameter,
      formType,
    } = this.state;

    const formIsDisabled = formType !== 2;

    const dropdownOptions: IDropdownOption[] = documentLocations.map(
      (location) => ({
        key: location.sharepointdocumentlocationid,
        text: location.name,
      })
    );
    const comboBoxOptions: IDropdownOption[] = documentLocations
      .filter((location) => location.parentsiteid)
      .map((location) => ({
        key: location.parentsiteid!,
        text: location.name,
      }));
    const files = this.getFilteredAndSortedFiles();
    const addIcon: IIconProps = { iconName: "Add" };
    const splitButtonStyles: IButtonStyles = {
      splitButtonMenuButton: {
        backgroundColor: "white",
        width: 28,
        border: "none",
      },
      splitButtonMenuIcon: { fontSize: "7px" },
      splitButtonDivider: {
        backgroundColor: "#c8c8c8",
        width: 1,
        right: 26,
        position: "absolute",
        top: 4,
        bottom: 4,
      },
      splitButtonContainer: {
        selectors: {
          [HighContrastSelector]: { border: "none" },
        },
      },
    };
    const menuProps: IContextualMenuProps = {
      items: [
        {
          key: "settings",
          iconProps: { iconName: "Settings" },
          text: "Settings",
          onClick: this.onGearIconClick,
        },
      ],
    };
    const isEmpty = sharePointDocLoc
      ? sharePointData.length === 0
      : files.length === 0;
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
                  data={previewFile.documentbody.replace(
                    /(data:application\/pdf;base64,).*?\1/,
                    "$1"
                  )}
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
              {isImage(previewFile.mimetype) && (
                <Img
                  src={previewFile.documentbody.replace(
                    /(data:image\/[a-zA-Z]+;base64,).*?\1/,
                    "$1"
                  )}
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

          <Dropzone onDrop={this.handleDrop} disabled={formIsDisabled}>
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
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          width: "100%",
                        }}
                      >
                        <div>
                          {currentFolderPath && (
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
                          )}
                        </div>
                        <div style={{ flexGrow: 1, textAlign: "center" }}>
                          {isEmpty ? (
                            <p>Drag and drop files here or Browse for files</p>
                          ) : (
                            this.renderFileList()
                          )}
                        </div>
                      </div>
                    </>
                  )}
                </div>
              </div>
            )}
          </Dropzone>
          {sharePointEnabledParameter && (
            <Stack
              horizontal
              style={{
                width: "100%",
                justifyContent: "space-between",
                alignItems: "center",
                flexWrap: "wrap",
              }}
            >
              <Toggle
                label={<div>SharePoint Documents {docLoctooltip}</div>}
                inlineLabel
                onText="On"
                offText="Off"
                checked={sharePointDocLoc}
                onChange={this.toggleSharePointDocLoc}
                disabled={sharePointEnabled}
                styles={{ root: { marginBottom: "0px" } }}
              />
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 5 }}
              >
                {isCalloutVisible && (
                  <Callout
                    target={this.targetRef.current}
                    onDismiss={this.onCalloutDismiss}
                    directionalHint={DirectionalHint.bottomAutoEdge}
                    setInitialFocus
                    role="dialog"
                  >
                    <Stack tokens={{ childrenGap: 10 }} padding={20}>
                      <Stack
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 10 }}
                      >
                        <Label>Remember Location?</Label>
                        <Toggle
                          onChange={this.toggleRememberLocation}
                          styles={{ root: { marginBottom: "0px" } }}
                          checked={userPreference}
                        />
                      </Stack>
                    </Stack>
                  </Callout>
                )}
                {sharePointDocLoc && (
                  <>
                    <Dropdown
                      options={dropdownOptions}
                      disabled={false}
                      styles={dropdownStyles}
                      onChange={this.handleDropdownChange}
                      selectedKey={selectedDocumentLocation}
                    />
                    {!formIsDisabled && (
                      <div ref={this.targetRef}>
                        <IconButton
                          split
                          iconProps={addIcon}
                          splitButtonAriaLabel="Location options"
                          aria-roledescription="Split button"
                          styles={splitButtonStyles}
                          menuProps={menuProps}
                          ariaLabel="New item"
                          onClick={this.openCreateLocationDialog}
                        />
                      </div>
                    )}
                  </>
                )}
                <Dialog
                  hidden={!isCreateLocationDialogVisible}
                  onDismiss={this.closeCreateLocationDialog}
                  dialogContentProps={{
                    type: DialogType.normal,
                    title: "Add Location",
                  }}
                >
                  <Stack tokens={{ childrenGap: 20 }} padding={20}>
                    <Label>
                      Create a new document location in Microsoft Dynamics 365
                    </Label>
                    <TextField
                      label="Display Name"
                      required
                      name="createLocationDisplayName"
                      value={createLocationDisplayName}
                      onChange={this.handleInputChange}
                    />
                    <Label>
                      Create new folder at the specified parent site
                    </Label>
                    <ComboBox
                      label="Parent Site"
                      required
                      selectedKey={selectedCreateLocation}
                      options={comboBoxOptions}
                      onChange={this.handleCreateLocationDropdownChange}
                    />
                    <TextField
                      label="Folder Name"
                      required
                      name="createLocationFolderName"
                      value={createLocationFolderName}
                      onChange={this.handleInputChange}
                    />
                  </Stack>
                  <DialogFooter>
                    <PrimaryButton
                      onClick={this.handlecreateLocation}
                      text="Save"
                      disabled={!isSaveButtonEnabled}
                    />
                    <DefaultButton
                      onClick={this.closeCreateLocationDialog}
                      text="Cancel"
                    />
                  </DialogFooter>
                </Dialog>
              </Stack>
            </Stack>
          )}
        </div>
      </>
    );
  }
}

export default Landing;
