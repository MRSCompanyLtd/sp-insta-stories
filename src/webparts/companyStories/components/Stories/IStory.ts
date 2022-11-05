export interface IStory {
    Id: number;
    Image: string;
    Author: {
        Id: number;
        Name: string;
        Email: string;
        Photo: string;
    };
    Content: string | undefined;
}