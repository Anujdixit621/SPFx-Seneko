import * as React from 'react';
import styles from './ApiCallingApp.module.scss';
import type { IApiCallingAppProps } from './IApiCallingAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, TextField, ChoiceGroup } from 'office-ui-fabric-react';
import { ISearchQuery, sp } from '@pnp/sp/presets/all';

// Declare global variables
const siteName = "https://pathinfotech365.sharepoint.com/sites/IT/";
const libraryName1 = "Path_Event";
interface DocTextResult {
  Title: string;
  FileRef: string;
  Author: string;
  Created: string;
  Modified: string;
}

export default class ApiCallingApp extends React.Component<IApiCallingAppProps, { radioValue: string, searchText: string, libraryName1:string }> {
  constructor(props: IApiCallingAppProps) {
    super(props);

    // Initial state
    this.state = {
      radioValue: 'allFields', // Default to 'All Fields'
      searchText: '', // Initially empty
      libraryName1:''
    };

    sp.setup({
      sp: { baseUrl: siteName },
    });
  }

  // Handle radio button change
  private onRadioChange = (_: any, option: any): void => {
    this.setState({ radioValue: option.key });
  };

  // Handle text input change
  private onTextChange = (_: any, newValue: string): void => {
    this.setState({ searchText: newValue });
  };
// Handle library input change 
private onLibrayNameInput = (_: any, newValue: string): void => {
  this.setState({ libraryName1: newValue });
};
  // Advanced search function
  private advancedSearchResult = async (): Promise<void> => {
    const { radioValue, searchText,libraryName1 } = this.state;

    try {
      const AddFilter = "KKS eq 'Yes'"; // Add fields by single value and key related to the library
      if (AddFilter) {
        const result1 = await sp.web.lists.getByTitle(libraryName1).items
          .select("Title", "FileLeafRef", "Author/Title", "FileRef")
          .expand("Author")
          .filter(AddFilter)
          .get();

        console.info("Result 1 by add field", result1);
      }

      if (searchText) {
        if (radioValue === 'documentText') {
          const result2 = this.freeTextSearch(searchText, libraryName1);
          console.info("Free text result 2 Document Text", result2);
        } else if (radioValue === 'entityName') {
          const result2 = this.entitySearch(searchText, libraryName1);
          console.info("Entity result 2", result2);
        } else if (radioValue === 'allFields') {
          const result2 = this.allfieldSearch(searchText, libraryName1);
          console.info("All fields result 2", result2);
        }
      }

      const contentType = "DocType";
      if (contentType) {
        const result3 = this.dataFilterByContentType(contentType, libraryName1);
        console.info("Filter Data By Content Type result 3", result3);
        const fieldsOfContetnType = this.getLibraryContentTypeFields(libraryName1, "DocType");
        if(fieldsOfContetnType){
          console.info(fieldsOfContetnType)
        }

      }
    } catch (error) {
      console.error("Error in advanced search:", error);
    }
  };
private async getLibraryContentTypeFields(libraryName: string, contentTypeName: string):Promise<void>{
try {
  const library = sp.web.lists.getByTitle(libraryName);
  if (library){
    const contentTypes = await library.contentTypes.filter(`Name eq '${contentTypeName}'`).get();
    const contentTypeId = contentTypes[0].StringId;
    if(contentTypes.length > 0){
      const fields = await library.contentTypes.getById(contentTypeId).fields.filter("Hidden eq false and ReadOnlyField eq false and FromBaseType eq false").get();
      console.log("Fields of the content type in the library:", fields);
    }else{
      console.log("Content Type not found in the library.");
    }
  } else{
    console.log("Library not found.");
  }
}
catch(e){
  console.error(e)
}
}
  
  // ... other functions (freeTextSearch, entitySearch, etc.) remain unchanged
  private freeTextSearch = async (searchText: string, libraryName: string): Promise<DocTextResult[]> => {
      
    const siteURL = "https://pathinfotech365.sharepoint.com/sites/IT"; // Provide the correct site URL
    const Path = `${siteURL}/${libraryName}/*`;
    console.info(`'${searchText} AND Path:${Path}'`);
    
    const searchQuery: ISearchQuery = {
      Querytext: `${searchText} AND Path:"${Path}"`,  // Ensure the search text includes the path
      RowLimit: 50,  // Number of search results
      SelectProperties: [
        "Title",          
        "Path",          
        "FileRef",        
        "Author",         
        "Created",        
        "Modified"       
      ],
      TrimDuplicates: true 
    };
  
    try {
      const searchResults = await sp.search(searchQuery);
      console.log("Search results:", searchResults.PrimarySearchResults);
  
      // Process the search results and map them to the desired format
      const finalArray: DocTextResult[] = searchResults.PrimarySearchResults.map((result: any) => ({
        Title: result.Title,
        FileRef: result.Path, // Adjust to result.Path or result.FileRef if needed
        Author: result.Author ? result.Author.Title : "Unknown",  // Handle cases where Author might be null
        Created: result.Created,
        Modified: result.Modified
      }));
      
      return finalArray;
    } catch (error) {
      console.error("Error performing search:", error);
      alert("Error performing search. Please try again.");
      return []; // Return an empty array in case of an error
    }
  };
  private entitySearch = async (searchText: string, libraryName: string):Promise<DocTextResult[]>=>{
    try{
      const siteName = "https://pathinfotech365.sharepoint.com/sites/IT/";
      sp.setup({
        sp: { baseUrl: siteName },
      });
      
      const AddFilter = `substringof('${searchText}', FileLeafRef)`;
      const entityResults=await sp.web.lists.getByTitle(libraryName).items
      .select("Title", "FileLeafRef", "Author/Title", "FileRef")
      .expand("Author")
      .filter(AddFilter) // Example filter: File name equals 'myfile.docx'
      .get(); 
      const finalArray: DocTextResult[] = entityResults.map((result: any) => ({
        Title: result.FileLeafRef,
        FileRef: result.FileRef, // Adjust to result.Path or result.FileRef if needed
        Author: result.Author ? result.Author.Title : "Unknown",  // Handle cases where Author might be null
        Created: result.Created,
        Modified: result.Modified
      }));

      return finalArray;
    } 
    catch(error){
    return error;
    }
  };
  private allfieldSearch = async(searchText: string, libraryName: string):Promise<void>=>{
    try{
         const library = sp.web.lists.getByTitle(libraryName);
         const fields = await library.fields.select("InternalName", "TypeAsString").filter("Hidden eq false").get();
         const TaxNomyFields =fields.filter(field => field.TypeAsString === "TaxonomyFieldType");	
				 const textFields = fields.filter(field => field.TypeAsString === "Text");
         const searchText1 ="Math";
				 const filterConditions = textFields.filter(field=>field.InternalName.indexOf("_")).map(field => `substringof('${searchText1}', ${field.InternalName})`).join(" or ");
         const Path_Event = "Path_Event";
				 const textFieldResults = await sp.web.lists.getByTitle(Path_Event).items.filter(filterConditions).get();
         console.info("textField result",textFieldResults)
				 let camlQuery = this._buildTaxonomyCAMLQuery(TaxNomyFields, searchText);
         let taxonomyFieldResults = [];
         if (camlQuery) {
          taxonomyFieldResults = await library.getItemsByCAMLQuery({
            ViewXml: camlQuery,
          });
        }
        const allResults = [...textFieldResults, ...taxonomyFieldResults];
        console.info(" allfieldSearch ",allResults);
			
       }
         catch(error){
    
       }
    };
  private _buildTaxonomyCAMLQuery = (fields: any[], searchText: string): string => {
      try {
          // Recursively build nested <Or> conditions
          const buildNestedOr = (conditions: string[]): string => {
              if (conditions.length === 1) {
                  return conditions[0]; // Return the last condition if only one left
              } else if (conditions.length === 2) {
                  return `<Or>${conditions[0]}${conditions[1]}</Or>`; // Wrap two conditions
              } else {
                  const firstCondition = conditions.shift(); // Remove the first condition
                  return `<Or>${firstCondition}${buildNestedOr(conditions)}</Or>`; // Recursively build <Or> conditions
              }
          };
  
          // Build conditions for each field dynamically
          const conditions = fields.map(field => 
              `<Eq><FieldRef Name='${field.InternalName}' /><Value Type='TaxonomyFieldType'>${searchText}</Value></Eq>`
          );
  
          // Return the full CAML query with proper <Or> structure
          return `<View><Query><Where>${buildNestedOr(conditions)}</Where></Query></View>`;
      } catch (error) {
          console.error("Error building CAML query:", error);
          return ""; // Handle the error and return an empty string
      }
    };

 private dataFilterByContentType = async (contentType:string, libraryName:string) =>{
      try{
        const items = await sp.web.lists
              .getByTitle(libraryName)
              .items.filter(`ContentType eq '${contentType}'`)
              .select("Title", "Id", "FileLeafRef", "FileRef", "ContentType/Name")
              .expand("ContentType")
              .get();
              return items;
      }
      catch(error){
        return [];
      }}


  public render(): React.ReactElement<IApiCallingAppProps> {
    
    const { description } = this.props;
    const { searchText, radioValue } = this.state;

    return (
      <section>
        <div className={styles.welcome}>
          <div>
         
          <TextField
              label="Search Text"
              value={libraryName1}
              onChange={this.onLibrayNameInput}
              placeholder="Enter keyword..."
            />
            <TextField
              label="Search Text"
              value={searchText}
              onChange={this.onTextChange}
              placeholder="Enter keyword..."
            />
            <ChoiceGroup
              label="Search By"
              selectedKey={radioValue}
              onChange={this.onRadioChange}
              options={[
                { key: 'documentText', text: 'Document Text' },
                { key: 'entityName', text: 'Entity Name' },
                { key: 'allFields', text: 'All Fields' }
              ]}
            />
            <PrimaryButton
              text="Search"
              onClick={this.advancedSearchResult}
            />
          </div>

          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>

        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams, and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting, and industry-standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
        </div>
      </section>
    );
  }
}
