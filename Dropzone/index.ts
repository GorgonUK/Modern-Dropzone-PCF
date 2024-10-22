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
    console.log("2.6");
    this.theContainer = container;

    if (typeof Xrm !== "undefined") {
      this.previousFormType = Xrm.Page.ui.getFormType() as number;
      const formContext = Xrm?.Page;
      if (formContext) {
        formContext.data.entity.addOnSave(this.checkFormTypeChange);
      }
    } else {
      console.error(
        "Xrm is not defined. Ensure this is run within a Dynamics 365 form."
      );
    }
    this.notifyOutputChanged = notifyOutputChanged;
    this.webAPI = context.webAPI;
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
    return React.createElement(Landing, { context: context });
  }
  public getOutputs(): IOutputs {
    return {};
  }
  public destroy(): void {}
}
