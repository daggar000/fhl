import React from "react";
import ReactDOM from "react-dom";
import App from "./components/App";
import "./index.css";
import { Provider, teamsTheme } from '@fluentui/react-northstar'
ReactDOM.render(
    <Provider theme={teamsTheme}>
        <App />
    </Provider>,
    document.getElementById("root"));
