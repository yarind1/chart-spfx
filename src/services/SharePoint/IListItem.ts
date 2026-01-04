export interface IListItem {
    Id: number;
    Year: number;
    [index: string]: any;
}

export  interface IListItemCollection {
    value: IListItem[];
}