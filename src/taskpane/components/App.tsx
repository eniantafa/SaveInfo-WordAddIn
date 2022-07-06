/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prettier/prettier */
import * as React from "react";
import { DefaultButton, TextField } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import $ from "jquery";
import { useMsal } from "@azure/msal-react";
import { callMsGraph } from "./graph";

import { loginRequest } from "./authConfig";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  Name: string;
  Job: string;
  Department: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      Name: "",
      Job: "",
      Department: "",
    };
    this.onchangedName = this.onchangedName.bind(this);
    this.onchangedJob = this.onchangedJob.bind(this);
    this.onchangedDepartment = this.onchangedDepartment.bind(this);
  }
  private onchangedName(name: any): void {
    this.setState({ Name: name.target.value });
  }
  private onchangedJob(job: any): void {
    this.setState({ Job: job.target.value });
  }
  private onchangedDepartment(department: any): void {
    this.setState({ Department: department.target.value });
  }
  componentDidMount() {
    // this.RequestProfileData();
   
  }

 

  click = async () => {
    // eslint-disable-next-line no-undef
    console.log(this.state.Name)
    // eslint-disable-next-line no-undef
    console.log(this.state.Department)
    // eslint-disable-next-line no-undef
    console.log(this.state.Job)


    $.ajax({
      async: true,
      crossDomain: true,
      url:
        "https://howling-crypt-47129.herokuapp.com/https://login.microsoftonline.com/" +
        "xhubnet.com" +
        "/oauth2/v2.0/token", // Pass your tenant name instead of sharepointtechie
      method: "POST",
      headers: {
        "access-control-allow-origin": "*",
      },
      data: {
        grant_type: "client_credentials",
        "client_id ": "5e0056b9-665b-421f-a95d-86533f14327b", //Provide your app id
        client_secret: "oDC8Q~Q.ae78fjWD6ORyVru6OTDDX_hyPDqDGdsO", //Provide your secret
        "scope ": "https://graph.microsoft.com/.default",
      },
      success: function (response) {
        var token = response.access_token;
        // eslint-disable-next-line no-undef
        console.log(token);
         //Fetch the values from the input elements  
    
  
    $.ajax({  
        async: true, // Async by default is set to “true” load the script asynchronously  
        // URL to post data into sharepoint list  
        url: "https://xhubnet.sharepoint.com/sites/EnianTafa" + "/_api/web/lists/GetByTitle('Profile')/items",  
        method: "POST", //Specifies the operation to create the list item  
        data: JSON.stringify({  
            '__metadata': {  
                'type': 'SP.Data.ProfileListItem' // it defines the ListEnitityTypeName  
            },  
//Pass the parameters
            'Name': this.state.Name,  
            'Job': this.state.Job,  
            'Department': this.state.Department,  
             
        }),  
        headers: {  
          Authorization: "Bearer " + token,
            "accept": "application/json;odata=verbose", //It defines the Data format   
            "content-type": "application/json;odata=verbose", //It defines the content type as JSON  
        },  
        success: function() {  
          // eslint-disable-next-line no-undef
          console.log("Item created successfully", "success"); // Used sweet alert for success message  
        },  
        error: function(error) {  
            // eslint-disable-next-line no-undef
            console.log(JSON.stringify(error));  
  
        }  
  
    })
        // $.ajax({
        //   method: "GET",
        //   url:
        //     "https://graph.microsoft.com/v1.0/sites/" +
        //     "5b90ebc3-f706-4b81-9ac6-6f162a22f23b" +
        //     "/lists/" +
        //     "%7Bd32024ab-d76f-48a8-9dbb-74cbb3da2ffb%7D" +
        //     "/items?expand=fields(select=Title,Name)')",
        //   headers: {
        //     Authorization: "Bearer " + token,
        //     "Content-Type": "application/json",
        //     "X-Requested-With": "XMLHttpRequest",
        //   },
        //   success: function (response) {
        //     var data = response.value;
        //     // eslint-disable-next-line no-undef
        //     console.log(data);
  
        //     // if(data.length==0){
  
        //     // }else{
        //     // debugger
        //     //   document.getElementById("run").style.display="none";
        //     //   document.getElementById("DIVNessunPostoLibero").style.display="";
        //     //   document.getElementById("PrenotatoManierExclusiva").style.display="";
  
        //     //   for(var i=0;i<=data.length-1;i++){
  
        //     //     Uteneti+=data[i].fields.Title+" - "+data[i].fields.TblPrenotazioni_Postazione+"\n";
  
        //     // }
        //     // document.getElementById('ListaUtenti').innerText=Uteneti;
        //     // }
        //   },
        // });
      },
    });
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      //const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      //paragraph.font.color = "blue";

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <div className="form-group">
          <TextField style={{ zIndex: 1000 }} label="Name" defaultValue="" onChange={this.onchangedName} />
        </div>
        <div className="form-group">
          <TextField style={{ zIndex: 1000 }} label="Job" defaultValue="" onChange={this.onchangedJob} />
        </div>
        <div className="form-group">
          <TextField style={{ zIndex: 1000 }} label="Department" defaultValue="" onChange={this.onchangedDepartment} />
        </div>
        <DefaultButton
          style={{ marginTop: "20px" }}
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.click}
        >
          Save
        </DefaultButton>
      </div>
      
    );
  }
}
