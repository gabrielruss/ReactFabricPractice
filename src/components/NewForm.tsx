/// <reference path="../../typings/SharePoint.d.ts" />
import * as React from "react";
import { Promise } from "es6-promise";

import { Button, ButtonType, Label } from '../../node_modules/office-ui-fabric-react/lib/index';

import { MyTextField } from './MyTextField';
import { PeoplePickerExample } from './MyPeoplePicker';

import { RestUtil } from '../../utils/RestUtil';

export class NewForm extends React.Component<any, any> {
    constructor(props: any) {
        super(props);

        this._handleChanged = this._handleChanged.bind(this);
        this._onSave = this._onSave.bind(this);

        this.state = { columns: {}, fieldsWithErrors: {} };
    }

    private _onSave(event: any) {
        event.preventDefault();

        RestUtil.getListData('ListItemEntityTypeFullName')
            .then((listData) => {
                return Promise.all([listData, RestUtil.submit(this.state.columns, listData)])
            })
            .then((response) => {
                alert("success");
            }, (error: any) => {
                alert(`There was a problem submitting your request: ${error}`);
            });
    }

    private _handleChanged(name: string, value: any, errorThrown?: string) {
        if (errorThrown) {
            this.state.fieldsWithErrors[`${name}`] = errorThrown;
            console.log("Error fields");
            console.log(this.state.fieldsWithErrors);
        } else {
            delete this.state.fieldsWithErrors[`${name}`];
        }
        this.state.columns[`${name}`] = value;
        console.log("Current State");
        console.log(this.state.columns);
        this.setState({
            columns: this.state.columns,
            fieldsWithErrors: this.state.fieldsWithErrors
        });
    }

    public render() {
        return (
            <div>
                <MyTextField
                    label="Title"
                    name="Title"
                    required={true}
                    onChanged={this._handleChanged} /> 
                {/* NEED TO PUT IN LOGIC TO HANDLE MULTIPLE USER FIELD VS SINGLE USER FIELD */}
                <PeoplePickerExample 
                    label="People"
                    name="People"
                    onChange={this._handleChanged}
                />
                <br/>   
                <Button
                    onClick={this._onSave}
                    type={"submit"}
                    buttonType={ButtonType.primary}
                    disabled={Object.keys(this.state.fieldsWithErrors).length !== 0}>
                    Save
                </Button>
            </div>
        );
    }
}