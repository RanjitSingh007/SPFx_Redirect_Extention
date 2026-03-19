declare interface IRedirectExtensionCommandSetStrings {
  NewItem: string;
  EditItem: string;
  ViewItem: string;
}

declare module 'RedirectExtensionCommandSetStrings' {
  const strings: IRedirectExtensionCommandSetStrings;
  export = strings;
}
