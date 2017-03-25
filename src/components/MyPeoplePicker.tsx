/* tslint:disable */
import * as React from 'react';
import { Promise } from "es6-promise";
/* tslint:enable */
// NEED TO CLEAN UP IMPORTS
import {
    BaseComponent,
    assign,
    autobind
} from '../../node_modules/office-ui-fabric-react/lib/Utilities';
import { IPersonaProps } from '../../node_modules/office-ui-fabric-react/lib/Persona';
import { IContextualMenuItem } from '../../node_modules/office-ui-fabric-react/lib/ContextualMenu';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    ListPeoplePicker,
    NormalPeoplePicker,
} from '../../node_modules/office-ui-fabric-react/lib/Pickers';
import { Button, ButtonType, Label } from '../../node_modules/office-ui-fabric-react/lib/index';
import { IPersonaWithMenu } from '../../node_modules/office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { RestUtil } from '../../utils/RestUtil';

// keeping this in for example state
export interface IPeoplePickerExampleState {
    delayResults?: boolean;
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};

export class PeoplePickerExample extends BaseComponent<any, IPeoplePickerExampleState> {
    private _peopleList;
    private contextualMenuItems: IContextualMenuItem[] = [
        {
            key: 'newItem',
            icon: 'circlePlus',
            name: 'New'
        },
        {
            key: 'upload',
            icon: 'upload',
            name: 'Upload'
        },
        {
            key: 'divider_1',
            name: '-',
        },
        {
            key: 'rename',
            name: 'Rename'
        },
        {
            key: 'properties',
            name: 'Properties'
        },
        {
            key: 'disabled',
            name: 'Disabled item',
            disabled: true
        }
    ];

    constructor(props: any) {
        super(props);
        this._handleChange = this._handleChange.bind(this);

        this._peopleList = [];
        this.state = {
            delayResults: false
            // create a state for filteredPersonas
        };
    }

    private _handleChange(items: IPersonaProps[], errorThrown?: string) {
        // instead of sending items, send the PersonId: 1
        // http://stackoverflow.com/questions/14556418/saving-a-person-or-group-field-using-rest
        this.props.onChange(this.props.name, items, errorThrown);
    }

    public render() {
        return (
            <div>
                <Label>{this.props.label}</Label>
                <CompactPeoplePicker
                    onChange={this._handleChange}
                    onResolveSuggestions={this._onFilterChanged}
                    getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
                    pickerSuggestionsProps={suggestionProps}
                    className={'ms-PeoplePicker'}
                />
            </div>
        );
    }

    @autobind
    private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        if (filterText) {
            let filteredPersonas:IPersonaProps[] = [];
            RestUtil.getUsers(filterText).then((results) => {
            for (let result in results) {
                filteredPersonas.push({
                    key: parseInt(result),
                    primaryText: results[result]["Name"]
                });
                console.log(`Found ${results[result]["Name"]}`);
            }
            filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
            filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
            // SET STATE FOR filteredPersonas???
        });
            return this._convertResultsToPromise(filteredPersonas);
        } else {
            return [];
        }
    }

    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }

    private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.primaryText === persona.primaryText).length > 0;
    }
    //need to use this function to delay the search results
    private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
        console.log(`_convertResultsToPromise results: ${results}`);
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }
}