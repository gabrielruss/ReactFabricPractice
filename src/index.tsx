/// <reference path="../typings/SharePoint.d.ts" />
import * as React from "react";
import * as ReactDOM from "react-dom";
import * as $ from "jquery";

import { NewForm } from "./components/NewForm";

$(document).ready(function(){
    ExecuteOrDelayUntilScriptLoaded(loaded, "sp.js");
});

function loaded() {
    ReactDOM.render(
        <NewForm />,
        document.getElementById("newform")
    );
}