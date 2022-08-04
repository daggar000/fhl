
import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import './App.css';
import { Button, Flex } from '@fluentui/react-northstar'
import axios from 'axios';
import './tab.css';

class Tab extends React.Component {
  constructor(props) {
    super(props)
    this.state = {
      context: {},
    }
  }
  //React lifecycle method that gets called once a component has finished mounting
  //Learn more: https://reactjs.org/docs/react-component.html#componentdidmount
  componentDidMount() {
    microsoftTeams.initialize();
    // Get the user context from Teams and set it in the state
    microsoftTeams.getContext((context, error) => {
      this.setState({
        context: context
      });
    });
    // Next steps: Error handling using the error object
  }



  render() {
    // this.publishAgenda();
    console.log(localStorage.getItem("key"));
    localStorage.setItem("key", "default");
    if (!localStorage.getItem("key"))
      localStorage.setItem("count", 0);
    return (
      <div align="left">
        <table width="250px">
          <tbody>
            {/* <tr>
              <td align="left">
                <img src="./images/teabreak1.png" alt="tea Break"></img>
              </td>
              <td>
                <img src="./images/meeting1.jfif" alt="Another meeting"></img>
              </td>
            </tr> */}
            <tr>
              <td align="center">
                To Take a short Break (5-10 mins).
              </td>
              <td align="center">
                Incase if you want to join another meeting.
              </td>
            </tr>
            <tr>
              <td align="center">
                <div id="buttons" stlye="float: left;">
                  <Flex gap="gap.smaller">
                    <Button onClick={this.beRightBack}>
                      Take a Break!
                    </Button>
                  </Flex>
                </div>
                {/* <label class="switch">
                  <input type="checkbox" />
                  <span class="slider round"></span>
                </label> */}
              </td>
              <td>
                <div id="buttons">
                  <Flex gap="gap.smaller">
                    <Button onClick={this.meetingSwitch}>
                      Another Meeting
                    </Button>
                  </Flex>
                  {/* <label class="switch">
                    <input type="checkbox"/>
                    <span class="slider round"></span>
                  </label> */}
                </div>
              </td>
            </tr>
            <tr >
              <td colSpan="2">
                <br></br>
                <h6>Below is the list of members who are on break:</h6>
                <hr></hr>
              </td>
            </tr>
            <tr >
              <td colSpan="2">
                <table id="brb-list" width="100%">

                </table>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    );
  }

  beRightBack = () => {
    var list2 = document.getElementById("brb-list")
    if (list2) {
      const userPrincipleName = this.state.context['userPrincipalName'] ?? "";
      this.syncApp(userPrincipleName, "TeaBreak").then(response => {
        const brbList = response.data;
        const personsObject = JSON.parse(brbList);
        let myMap = new Map(Object.entries(personsObject));
        this.populateUI(myMap);
      });
    }
  }
  meetingSwitch = () => {
    var list2 = document.getElementById("brb-list")
    if (list2) {
      const userPrincipleName = this.state.context['userPrincipalName'] ?? "";
      this.syncApp(userPrincipleName, "meetingSwitch").then(response => {
        const brbList = response.data;
        const personsObject = JSON.parse(brbList);
        let myMap = new Map(Object.entries(personsObject));
        this.populateUI(myMap);
      });
    }
  }

  populateUI(BRBMap) {
    var list2 = document.getElementById("brb-list");
    list2.style.marginTop = "50px";
    list2.innerHTML = '';
    BRBMap.forEach((value, key) => {
      list2.appendChild(this.createRow(key, value));
    });

  }

  createRow(BRBName, breakTypeValue) {
    const row = document.createElement("tr");
    const name = document.createElement("td");
    const breakType = document.createElement("td");

    name.innerHTML = BRBName;
    breakType.innerHTML = breakTypeValue;
    row.appendChild(name);
    row.appendChild(breakType);
    return row;
  }

  async syncApp(userPrincipleName, breakType) {

    const agendaValue = userPrincipleName.replace('@microsoft.com', '');
    var publishData = { name: agendaValue, breaktype: breakType };
    let config = {
      headers: {
        "Content-Type": "application/json",
      }
    }
    let url = window.location.origin;
    console.log("url=" + url);
   let reponse = await axios.post('http://localhost:3978/api/sendAgenda', publishData, config);
   //let reponse = await axios.post('https://helloworlddevb65bdfbot.azurewebsites.net/api/sendAgenda', publishData, config);
    return reponse;
  }
}
export default Tab;
