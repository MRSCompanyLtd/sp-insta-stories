export interface IStoryReturn {
    Id: number;
    Title: string;
    Content: string;
    Image: string;
    Author?: {
        Id: number;
        Title: string;
        EMail: string;
    }
    Author0: {
        Id: number;
        Title: string;
        EMail: string;
    }
}