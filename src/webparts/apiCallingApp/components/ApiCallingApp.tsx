// import * as React from 'react';
// import styles from './ApiCallingApp.module.scss';
// import type { IApiCallingAppProps } from './IApiCallingAppProps';
// import { escape } from '@microsoft/sp-lodash-subset';
// import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
// import { ISearchQuery, sp } from '@pnp/sp/presets/all';
// import {  TextField, ChoiceGroup } from 'office-ui-fabric-react';
// // npm i @pnp/sp@2.11.0

// // Declare global variables
// const siteName = "https://{domain}/sites/IT/";
// const libraryName = "Path_Event";



// export default class ApiCallingApp extends React.Component<IApiCallingAppProps, { radioValue: string, searchText: string }> {
  
//   constructor(props: IApiCallingAppProps) {
//     super(props);
//     // Initial state
//     this.state = {
//       radioValue: 'allFields', // Default to 'All Fields'
//       searchText: '' // Initially empty
//     };

//     sp.setup({
//       sp: { baseUrl: siteName },
//     });
//   }

//   public render(): React.ReactElement<IApiCallingAppProps> {
//     interface DocTextResult {  // Declare interface here
//       Title: string;
//       FileRef: string;
//       Author: string;
//       Created: string;
//       Modified: string;
//     }
//     const {
//       description,
      
//     } = this.props;


    
//     const advancedSearchResult = async () => {
//       try
//       { 
//           const libraryName = "Path_Event";
//           const  keyword ="HITACHI"; // comes from input text
//          // const  enityName ="Alstom";
//           const radioDocumentText = false; 
//           const radioEnitiyName = false;
//           const radioAllfields = true;
//           const contentType = "DocType";
//           const AddFilter ="KKS eq 'Yes'"; // Add fields by single value and key are related to library 
//           if(AddFilter){
//              const result1 = await sp.web.lists.getByTitle(libraryName).items
//              .select("Title", "FileLeafRef", "Author/Title", "FileRef")
//              .expand("Author")
//              .filter(AddFilter) // Example filter: File name equals 'myfile.docx'
//              .get();
        
//              let finalResults = result1;
//              console.info("result 1 by add field",finalResults)
//           }
//             if(keyword){
//               if(radioDocumentText){
//                const result2 = freeTextSearch(keyword,libraryName  );
//                console.info("Free text result 2 Document Text",result2);
//               } else if(radioEnitiyName){
//                  //const  = "FileName"
//                  const result2 = entitySearch(keyword,libraryName);
//                  console.info("entity result 2 ",result2);
//               }
//               else if(radioAllfields){
//                 const result2 = allfieldSearch(keyword,libraryName);
//                 console.info("All field radio button result 2 ",result2);
//               }
              
//             }
//             if(contentType){
//               const result3 = dataFilterByContentType(contentType,libraryName);
//               console.info("Filter Data By Content Type result3",result3)
//             }
//       } 

//       catch
//       {
// return[];
//       }
     
//     };
//     interface DocTextResult {
//       Title: string;
//       FileRef: string;
//       Author: string;
//       Created: string;
//       Modified: string;
//     }
    
//     const freeTextSearch = async (searchText: string, libraryName: string): Promise<DocTextResult[]> => {
      
//       const siteURL = "https://pathinfotech365.sharepoint.com/sites/IT"; // Provide the correct site URL
//       const Path = `${siteURL}/${libraryName}/*`;
//       console.info(`'${searchText} AND Path:${Path}'`);
      
//       const searchQuery: ISearchQuery = {
//         Querytext: `${searchText} AND Path:"${Path}"`,  // Ensure the search text includes the path
//         RowLimit: 50,  // Number of search results
//         SelectProperties: [
//           "Title",          
//           "Path",          
//           "FileRef",        
//           "Author",         
//           "Created",        
//           "Modified"       
//         ],
//         TrimDuplicates: true 
//       };
    
//       try {
//         const searchResults = await sp.search(searchQuery);
//         console.log("Search results:", searchResults.PrimarySearchResults);
    
//         // Process the search results and map them to the desired format
//         const finalArray: DocTextResult[] = searchResults.PrimarySearchResults.map((result: any) => ({
//           Title: result.Title,
//           FileRef: result.Path, // Adjust to result.Path or result.FileRef if needed
//           Author: result.Author ? result.Author.Title : "Unknown",  // Handle cases where Author might be null
//           Created: result.Created,
//           Modified: result.Modified
//         }));
        
//         return finalArray;
//       } catch (error) {
//         console.error("Error performing search:", error);
//         alert("Error performing search. Please try again.");
//         return []; // Return an empty array in case of an error
//       }
//     };
//     const entitySearch = async (searchText: string, libraryName: string):Promise<DocTextResult[]>=>{
//       try{
//         const siteName = "https://pathinfotech365.sharepoint.com/sites/IT/";
//         sp.setup({
//           sp: { baseUrl: siteName },
//         });
        
//         const AddFilter = `substringof('${searchText}', FileLeafRef)`;
//         const entityResults=await sp.web.lists.getByTitle(libraryName).items
//         .select("Title", "FileLeafRef", "Author/Title", "FileRef")
//         .expand("Author")
//         .filter(AddFilter) // Example filter: File name equals 'myfile.docx'
//         .get(); 
//         const finalArray: DocTextResult[] = entityResults.map((result: any) => ({
//           Title: result.FileLeafRef,
//           FileRef: result.FileRef, // Adjust to result.Path or result.FileRef if needed
//           Author: result.Author ? result.Author.Title : "Unknown",  // Handle cases where Author might be null
//           Created: result.Created,
//           Modified: result.Modified
//         }));

//         return finalArray;
//       } 
//       catch(error){
//       return error;
//       }
//     };
//     const allfieldSearch = async(searchText: string, libraryName: string):Promise<void>=>{
//     try{
//          const library = sp.web.lists.getByTitle(libraryName);
//          const fields = await library.fields.select("InternalName", "TypeAsString").filter("Hidden eq false").get();
//          const TaxNomyFields =fields.filter(field => field.TypeAsString === "TaxonomyFieldType");	
// 				 const textFields = fields.filter(field => field.TypeAsString === "Text");
//          const searchText1 ="Math";
// 				 const filterConditions = textFields.filter(field=>field.InternalName.indexOf("_")).map(field => `substringof('${searchText1}', ${field.InternalName})`).join(" or ");
//          const Path_Event = "Path_Event";
// 				 const textFieldResults = await sp.web.lists.getByTitle(Path_Event).items.filter(filterConditions).get();
//          console.info("textField result",textFieldResults)
// 				 let camlQuery = _buildTaxonomyCAMLQuery(TaxNomyFields, searchText);
//          let taxonomyFieldResults = [];
//          if (camlQuery) {
//           taxonomyFieldResults = await library.getItemsByCAMLQuery({
//             ViewXml: camlQuery,
//           });
//         }
//         const allResults = [...textFieldResults, ...taxonomyFieldResults];
//         console.info(" allfieldSearch ",allResults);
			
//        }
//          catch(error){
    
//        }
//     };
//     const _buildTaxonomyCAMLQuery = (fields: any[], searchText: string): string => {
//       try {
//           // Recursively build nested <Or> conditions
//           const buildNestedOr = (conditions: string[]): string => {
//               if (conditions.length === 1) {
//                   return conditions[0]; // Return the last condition if only one left
//               } else if (conditions.length === 2) {
//                   return `<Or>${conditions[0]}${conditions[1]}</Or>`; // Wrap two conditions
//               } else {
//                   const firstCondition = conditions.shift(); // Remove the first condition
//                   return `<Or>${firstCondition}${buildNestedOr(conditions)}</Or>`; // Recursively build <Or> conditions
//               }
//           };
  
//           // Build conditions for each field dynamically
//           const conditions = fields.map(field => 
//               `<Eq><FieldRef Name='${field.InternalName}' /><Value Type='TaxonomyFieldType'>${searchText}</Value></Eq>`
//           );
  
//           // Return the full CAML query with proper <Or> structure
//           return `<View><Query><Where>${buildNestedOr(conditions)}</Where></Query></View>`;
//       } catch (error) {
//           console.error("Error building CAML query:", error);
//           return ""; // Handle the error and return an empty string
//       }
//     };
//     const dataFilterByContentType = async (contentType:string, libraryName:string) =>{
//   try{
//     const items = await sp.web.lists
//           .getByTitle(libraryName)
//           .items.filter(`ContentType eq '${contentType}'`)
//           .select("Title", "Id", "FileLeafRef", "FileRef", "ContentType/Name")
//           .expand("ContentType")
//           .get();
//           return items;
//   }
//   catch(error){
//     return [];
//   }
//     }
//     // const _buildTaxonomyCAMLQuery1 = (fields: any[], searchText: string): string => {
//     //   try {
//     //       // Recursively build nested <Or> conditions
//     //       const buildNestedOr = (conditions: string[]): string => {
//     //           if (conditions.length === 1) {
//     //               return conditions[0]; // Return the last condition if only one left
//     //           } else if (conditions.length === 2) {
//     //               return `<Or>${conditions[0]}${conditions[1]}</Or>`; // Wrap two conditions
//     //           } else {
//     //               const firstCondition = conditions.shift(); // Remove the first condition
//     //               return `<Or>${firstCondition}${buildNestedOr(conditions)}</Or>`; // Recursively build <Or> conditions
//     //           }
//     //       };
//     //       const conditions = fields.map(field => 
//     //           `<Eq><FieldRef Name='${field.InternalName}' /><Value Type='TaxonomyFieldType'>${searchText}</Value></Eq>`
//     //       );
//     //       return `<View><Query><Where>${buildNestedOr(conditions)}</Where></Query></View>`;
//     //   } catch (error) {
//     //       console.error("Error building CAML query:", error);
//     //       return ""; 
//     //   }
  
//     // };
  
//     return (
//       <section >
//         <div className={styles.welcome}>
//         <div>
//         <TextField
//               label="Search Text"
//               value={searchText}
//               onChange={this.onTextChange}
//               placeholder="Enter keyword..."
//             />
//             <ChoiceGroup
//               label="Search By"
//               selectedKey={radioValue}
//               onChange={this.onRadioChange}
//               options={[
//                 { key: 'documentText', text: 'Document Text' },
//                 { key: 'entityName', text: 'Entity Name' },
//                 { key: 'allFields', text: 'All Fields' }
//               ]}
//             />
//       <PrimaryButton
//         text="Postman"
//         onClick={advancedSearchResult} // Call showAlert function on button click
//       />
//     </div>
        
//           <div>Web part property value: <strong>{escape(description)}</strong></div>
//         </div>
//         <div>
//           <h3>Welcome to SharePoint Framework!</h3>
//           <p>
//             The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//           </p>
//           <h4>Learn more about SPFx development:</h4>
//           <ul className={styles.links}>
          
//           </ul>
//         </div>
//       </section>
//     );
//   }
//   private onRadioChange = (_: any, option: any): void => {
//     this.setState({ radioValue: option.key });
//   };

//   // Handle text input change
//   private onTextChange = (_: any, newValue: string): void => {
//     this.setState({ searchText: newValue });
//   };
// }


import * as React from 'react';
import styles from './ApiCallingApp.module.scss';
import type { IApiCallingAppProps } from './IApiCallingAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton, TextField, ChoiceGroup } from 'office-ui-fabric-react';
import {   ISearchQuery, sp } from '@pnp/sp/presets/all';
import { ContentType } from '../model/ContentType';

// Declare global variables
const siteName = "https://pathinfotech365.sharepoint.com/sites/IT/";
//let libraryName = "";
interface DocTextResult {
  Title: string;
  FileRef: string;
  Author: string;
  Created: string;
  Modified: string;
}

export default class ApiCallingApp extends React.Component<IApiCallingAppProps, { radioValue: string, searchText: string, libraryName1:string}> {
  constructor(props: IApiCallingAppProps) {
    super(props);

    // Initial state
    this.state = {
      radioValue: 'allFields', // Default to 'All Fields'
      searchText: '',
      libraryName1:'' // Initially empty
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


  //onLibraryTextChange
  private onLibraryTextChange = (_: any, newValue: string): void => {
    this.setState({ libraryName1: newValue });
  };
  // Advanced search function
  
  private advancedSearchResult = async (): Promise<void> => {
    const { radioValue, searchText,libraryName1 } = this.state;

    try {
      const results = [];
      const AddFilter = "KKS eq 'Yes'"; // Add fields by single value and key related to the library
      if (AddFilter) {
        const result1 = await sp.web.lists.getByTitle(libraryName1).items
          .select("ID", "Title", "FileLeafRef", "Author/Title", "FileRef")
          .expand("Author")
          .filter(AddFilter)
          .get();
          results.push(...result1);
        console.info("Result 1 by add field", result1);
      }

      if (searchText) {
        if (radioValue === 'documentText') {
          const result2 = await this.freeTextSearch(searchText, libraryName1);
          results.push(... result2);
          console.info("Free text result 2 Document Text", result2);
        } else if (radioValue === 'entityName') {
          const result2 = await this.entitySearch(searchText, libraryName1);
          results.push(... result2);
          console.info("Entity result 2", result2);
        } else if (radioValue === 'allFields') {
          const result2 = await this.allfieldSearch(searchText, libraryName1);
          results.push(... result2);
          console.info("All fields result 2", result2);
        }
      }

      const contentType = "DocType"; // get from dropdown
      if (contentType) {
        const result3 = await this.dataFilterByContentType(contentType, libraryName1);
        results.push(... result3);
        console.info("Filter Data By Content Type result 3", result3);
        const fieldsOfContetnType = this.getLibraryContentTypeFields(libraryName1, "DocType");
        if(fieldsOfContetnType){
          console.info(fieldsOfContetnType)

          const result4 = await this.filterByContentTypeField(libraryName1);
          
          console.info("filterByContentTypeField ....",result4)
        }
        console.info("Final combined result:", results);
        const uniqueResults = this.removeDuplicatesByFileRef(results);
        console.log('Unique Result duplicacy by FileRef',uniqueResults)

      }
    } catch (error) {
      console.error("Error in advanced search:", error);
    }
  };

  private async getLibraryContentTypeFields(libraryName: string, contentTypeName: string): Promise<ContentType[]> {
    try {
      const library = sp.web.lists.getByTitle(libraryName);
      if (library) {
        const contentTypes = await library.contentTypes.filter(`Name eq '${contentTypeName}'`).get();
  
        if (contentTypes.length > 0) {
          const contentTypeId = contentTypes[0].StringId;
  
          const fields = await library.contentTypes
            .getById(contentTypeId)
            .fields.select("InternalName", "TypeAsString", "Id", "SchemaXml")
            .filter("Hidden eq false and ReadOnlyField eq false and FromBaseType eq false")
            .get();
  
          const contentTypeFields: ContentType[] = fields.map((field: any) => {
            return new ContentType(
              field.InternalName,
              field.TypeAsString
            );
          });
  
          // Log and return the ContentType fields
          console.log("Fields of the content type in the library:", contentTypeFields);
          return contentTypeFields;
        } else {
          console.log(`Content Type '${contentTypeName}' not found in the library.`);
          return [];  // Return empty array if content type is not found
        }
      } else {
        console.log(`Library '${libraryName}' not found.`);
        return [];  // Return empty array if library is not found
      }
    } catch (e) {
      console.error("Error retrieving content type fields:", e);
      return [];  // Return empty array in case of error
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
        "ID",          
        "Title",          
        "FileRef",        
        "Author/Title",         
        "FileRef"       
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
      .select("ID", "Title", "FileLeafRef", "Author/Title", "FileRef")
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
  private allfieldSearch = async(searchText: string, libraryName: string):Promise<any>=>{
    try{
         const library = sp.web.lists.getByTitle(libraryName);
         const fields = await library.fields.select("InternalName", "TypeAsString").filter("Hidden eq false").get();
         const TaxNomyFields =fields.filter(field => field.TypeAsString === "TaxonomyFieldType");	
				 const textFields = fields.filter(field => field.TypeAsString === "Text");
         const searchText1 ="Math";
				 const filterConditions = textFields.filter(field=>field.InternalName.indexOf("_")).map(field => `substringof('${searchText1}', ${field.InternalName})`).join(" or ");
         const Path_Event = "Path_Event";
				 const textFieldResults = await sp.web.lists.getByTitle(Path_Event).items.filter(filterConditions).select("ID", "Title", "FileLeafRef", "Author/Title", "FileRef").get();
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
              .select("ID", "Title", "FileLeafRef", "Author/Title", "FileRef")
              .expand("Author")
              .get();
              return items;
      }
      catch(error){
        return [];
      }}

      private filterByContentTypeField = async (libraryName: string) => {
        try {
          // Get the content type fields for the given library name and "DocType" content type
          const fieldsOfContetnType = await this.getLibraryContentTypeFields(libraryName, "DocType");
      
          // Filter the fields by type
          const fieldsOfText = fieldsOfContetnType.filter((field: { TypeAsString: string }) => field.TypeAsString === "Text");
          const fieldsOfTaxo = fieldsOfContetnType.filter((field: { TypeAsString: string }) => field.TypeAsString === "TaxonomyFieldType");
      
          // Log the filtered fields of type "Text"
          if (fieldsOfText.length > 0) {
            const filterConditions = [];
            for (const field of fieldsOfText) {
              const searchValue ="Path";
              const internalName = field.InternalName;
              filterConditions.push(`substringof('${searchValue}', ${internalName})`);
            }
            const filterQuery = filterConditions.join(' or ');
            console.info("filterQuery of content type form",filterQuery)
            const files = sp.web.lists.getByTitle(libraryName).items.select("ID", "Title", "FileLeafRef", "Author/Title", "FileRef").expand("Author")
            .filter(filterQuery)
            //filter("substringof('Path', Plant) or substringof('Path', DOC_x0020_No) or substringof('Path', Ref) or substringof('Path', Remark) or substringof('Path', KKS)")
            .get();
      
            console.info("Filter files by content type field:", files);
            return files;
          }
      
          // Log the filtered fields of type "TaxonomyFieldType"
          if (fieldsOfTaxo.length > 0) {
            console.info("Fields of type 'TaxonomyFieldType':", fieldsOfTaxo);
          }
         
        } catch (e) {
          console.error("Error in filtering by content type field:", e);
        }
      };

      private removeDuplicatesByFileRef = (data: any[]): any[] => {
        const seen = new Set();
        return data.filter(item => {
            const fileRef = item.FileRef; // Adjust based on your item structure
            if (seen.has(fileRef)) {
                return false; // This is a duplicate
            }
            seen.add(fileRef); // Add to seen set
            return true; // Keep this item
        });
    };
      

  public render(): React.ReactElement<IApiCallingAppProps> {
    
    const { description } = this.props;
    const { searchText, radioValue,libraryName1 } = this.state;

    return (
      <section>
        <div className={styles.welcome}>
          <div>
          <TextField
              label="Library Name"
              value={libraryName1}
              onChange={this.onLibraryTextChange}
              placeholder="Enter library Name..."
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
              text="PostMan"
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
