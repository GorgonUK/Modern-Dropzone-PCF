import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Landing } from "./Landing/Landing";
import * as React from "react";

export class Dropzone
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private theContainer: HTMLDivElement;
  private notifyOutputChanged: () => void;
  private webAPI: ComponentFramework.WebApi;
  private previousFormType: number | null = null;
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    console.log("Modern Dropzone PCF 2.9.3 Initialised");
    this.theContainer = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.webAPI = context.webAPI;
  
    const xrmCheckInterval = setInterval(() => {
      if (typeof Xrm !== "undefined") {
        clearInterval(xrmCheckInterval);
  
        this.previousFormType = Xrm.Page.ui.getFormType() as number;
        const formContext = Xrm?.Page;
  
        if (formContext) {
          formContext.data.entity.addOnSave(this.checkFormTypeChange);
        }
      } 
    }, 500);
  
    setTimeout(() => {
      clearInterval(xrmCheckInterval);
    }, 10000);
  }

  private checkFormTypeChange = (): void => {
    const formType = Xrm?.Page.ui.getFormType();
    if (formType !== 1) return;
    this.notifyOutputChanged();

    const pollForId = setInterval(() => {
      let entityId = Xrm?.Page.data.entity.getId();

      if (entityId) {
        clearInterval(pollForId);
        entityId = entityId.replace(/[{}]/g, "").toLowerCase();
        const saveEvent = new CustomEvent("recordSavedEvent", {
          detail: {
            entityId: entityId,
            timestamp: Date.now(),
          },
        });
        window.dispatchEvent(saveEvent);
      }
    }, 500);
  };
  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    let isDisabled = context.mode.isControlDisabled;
    return React.createElement(Landing, { context: context, isDisabled });
  }
  public getOutputs(): IOutputs {
    return {};
  }
  public destroy(): void {}
}
