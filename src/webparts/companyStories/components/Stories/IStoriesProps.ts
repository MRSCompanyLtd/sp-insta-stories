import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IStory } from "./IStory";

export interface IStoriesProps {
    stories: IStory[];
    context: WebPartContext;
    timing: number;
}