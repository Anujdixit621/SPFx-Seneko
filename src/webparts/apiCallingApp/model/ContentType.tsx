export class ContentType {
    InternalName: string;
    TypeAsString: string;
    editLink?: string;
    id?: string;
    type?: string;
  
    constructor(internalName: string, typeAsString: string) {
      this.InternalName = internalName;
      this.TypeAsString = typeAsString;
    }
  }
  