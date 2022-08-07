
import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import './App.css';
import { Button, Flex } from '@fluentui/react-northstar'
import axios from 'axios';
import './tab.css';
import { Input } from '@fluentui/react-northstar'
import { Divider } from '@fluentui/react-northstar'
import { CallVideoIcon, TeamCreateIcon, AccessibilityIcon, CallPstnIcon, GlassesIcon, LightningIcon, CustomerHubIcon, DoorArrowLeftIcon } from '@fluentui/react-icons-northstar'

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
        <table width="200px" >
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
              <td align="left">
                <div id="buttons" stlye="float: left;">
                  <Button onClick={this.beRightBack} icon={<DoorArrowLeftIcon />} text content="Break!" />
                </div>

              </td>
              <td align="left">
                <div id="buttons" stlye="float: left;">

                  <Button onClick={this.meetingSwitch} icon={<CallPstnIcon />} text content="Another Meeting!" />
                </div>
              </td>
            </tr>
            <tr>
              <td align="left">
                <div id="buttons" stlye="float: left;">
                  <Button onClick={this.customReason} icon={<CustomerHubIcon />} text content="Custom" />
                </div>
              </td>
              <td >
                <div id="buttons" stlye="float: left;" >
                  <form action="/action_page.php">
                    <Input type="text" id="fname" placeholder="your message?"></Input>
                  </form>
                </div>
              </td>
            </tr>
            <tr >
              <td colSpan="2">
                <br></br>
                <Divider content="List of members who are AFK:" />
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

  customReason = () => {
    var list2 = document.getElementById("brb-list")
    if (list2) {
      var customReason = document.getElementById("fname").value;
      const userPrincipleName = this.state.context['userPrincipalName'] ?? "";
      const meetingID = this.state.context['meetingId'] ?? "";
      this.syncFunction(userPrincipleName, "custom", meetingID, customReason);
    }
  }
  beRightBack = () => {
    var list2 = document.getElementById("brb-list")
    if (list2) {
      const userPrincipleName = this.state.context['userPrincipalName'] ?? "";
      const meetingID = this.state.context['meetingId'] ?? "";
      this.syncFunction(userPrincipleName, "teaBreak", meetingID, "");
    }
  }
  meetingSwitch = () => {
    var list2 = document.getElementById("brb-list")
    if (list2) {
      const userPrincipleName = this.state.context['userPrincipalName'] ?? "";
      const meetingID = this.state.context['meetingId'] ?? "";
      this.syncFunction(userPrincipleName, "meetingSwitch", meetingID, "");
    }
  }
  syncFunction(userPrincipleName, breakType, meetingID, desc) {
    this.syncApp(userPrincipleName, breakType, meetingID, desc).then(response => {
      const brbList = response.data;
      let myMap = new Map(Object.entries(brbList));
      this.populateUI(myMap);
    });
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
    const breakTypeEle = document.createElement("td");

    name.innerHTML = BRBName;
    breakTypeEle.innerHTML = this.getUIBreakContent(breakTypeValue);
    row.appendChild(name);
    row.appendChild(breakTypeEle);
    return row;
  }

  getUIBreakContent(breakTypeValue) {
    const breakType = breakTypeValue.split("::");
    if (breakType[0] === 'custom') {
      return breakType[1];
    }
    else if (breakType[0] === 'meetingSwitch') {
      return "has joined another meeting";
    }
    else {
      return "is out for small break!";
    }
  }
  
  async syncApp(userPrincipleName, breakTypeValue, meetingID, descr) {

    const agendaValue = userPrincipleName.replace('@microsoft.com', '');
    var publishData = { name: agendaValue, breaktype: breakTypeValue, meetingid: meetingID, desc: descr };
    let config = {
      headers: {
        "Content-Type": "application/json",
      }
    }
    //let reponse = await axios.post('http://localhost:7071/api/Function1', publishData, config);
    let reponse = await axios.post('https://timebreakapp20220806210728.azurewebsites.net/api/Function1', publishData, config);

    console.log(reponse.data);
    return reponse;
  }
}
export default Tab;
