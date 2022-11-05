import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { GraphFI } from "@pnp/graph";

export interface ICompanyStoriesProps {
  title: string;
  timing: string;
  sp: SPFI;
  graph: GraphFI;
  selectedList: string;
  context: WebPartContext;
}
