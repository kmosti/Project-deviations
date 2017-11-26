declare interface IProsjektAvvikWebPartStrings {
  PropertyPaneDescription: string;
  TemplateGroupName: string;
  TitleFieldLabel: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  FieldsMaxResults: string;
  PowerAppLink: string;

  /* Dialog */
  ScriptsDialogHeader: string;
  ScriptsDialogSubText: string;
}

declare module 'ProsjektAvvikWebPartStrings' {
  const strings: IProsjektAvvikWebPartStrings;
  export = strings;
}
