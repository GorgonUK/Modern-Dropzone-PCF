import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Landing } from "./Landing/Landing";
import * as React from "react";

export class Dropzone implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theContainer: HTMLDivElement;
    private notifyOutputChanged: () => void;
    private webAPI: ComponentFramework.WebApi;
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
      console.log("2.2");
        this.theContainer = container;
        this.notifyOutputChanged = notifyOutputChanged;
        this.webAPI = context.webAPI;
    }
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        return React.createElement(Landing, { context: context });
    }
    public getOutputs(): IOutputs {
        return {};
    }
    public destroy(): void {
    }
}
