import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { IRefinerProperties } from "../RefinersCard/IRefinerProperties";

export interface IRefinersCardProps {
  context: WebPartContext | ApplicationCustomizerContext;
  refinerProperties: IRefinerProperties;
}
