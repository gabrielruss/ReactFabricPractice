import * as React from 'react';
import { TextField } from '../../node_modules/office-ui-fabric-react/lib/index';

export class MyTextField extends React.Component<any, any> {

    constructor(props: any) {
        super(props);
        this._handleChange = this._handleChange.bind(this);
        this._getErrorMessage = this._getErrorMessage.bind(this);
    }

    private _handleChange(value: string, errorThrown?: string) {
        this.props.onChanged(this.props.name, value, errorThrown);
    }

    private _getErrorMessage(value: string): string {
        let errorMessage = '';
        if (this.props.required && value.length < 1) {
            errorMessage = 'Field is required.';
            this._handleChange(value, errorMessage);
        } else {
            this._handleChange(value, errorMessage);
        }
        return errorMessage;
    }

    public render() {
        return (
            <div>
                <TextField
                    className={this.props.className}
                    label={this.props.label}
                    required={this.props.required}
                    onChanged={this._handleChange}
                    onGetErrorMessage={this._getErrorMessage}
                    multiline={this.props.multiline}
                    placeholder={this.props.placeholder}
                    />
            </div>
        );
    }
}