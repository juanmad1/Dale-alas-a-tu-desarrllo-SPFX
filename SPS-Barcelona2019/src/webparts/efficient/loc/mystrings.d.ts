declare interface IEfficientWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  EmptyUrlMessage: string;
  ErroMessage: string;
  DataSourceLabel: string;
  Loading: string
}

declare module 'EfficientWebPartStrings' {
  const strings: IEfficientWebPartStrings;
  export = strings;
}
