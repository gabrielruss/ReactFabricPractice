import * as React from 'react';
import { Promise } from "es6-promise";
// NEED TO CLEAN UP IMPORTS
import {
    BaseComponent,
    autobind
} from '../../node_modules/office-ui-fabric-react/lib/Utilities';
import { IPersonaProps } from '../../node_modules/office-ui-fabric-react/lib/Persona';
import { IContextualMenuItem } from '../../node_modules/office-ui-fabric-react/lib/ContextualMenu';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
} from '../../node_modules/office-ui-fabric-react/lib/Pickers';
import { Label } from '../../node_modules/office-ui-fabric-react/lib/index';
import { RestUtil } from '../../utils/RestUtil';

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};

export class PeoplePickerExample extends BaseComponent<any, any> {
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
        this.state = {
            delayResults: false,
            userIds: {
                results: []
            }
        };
    }

    private _handleChange(items: IPersonaProps[], errorThrown?: string) {
        let tempArray = [];
        for (let item in items) {
            tempArray.push(items[item]["id"]);
        }
        this.state.userIds.results = tempArray;
        this.props.onChange(`${this.props.name}Id`, this.state.userIds, errorThrown);
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
                    inputProps={{ disabled: this.state.userIds.results.length >= 1 && !this.props.multipleUsers}}
                />
                {/* make the inputProps thingy optional based on props passed from parent */}
                {/**/}
            </div>
        );
    }

    @autobind
    private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        return new Promise((resolve, reject) => {
            if (filterText) {
                let filteredPersonas: IPersonaProps[] = [];
                RestUtil.getUsers(filterText).then((results) => {
                    for (let result in results) {
                        filteredPersonas.push({
                            key: parseInt(result),
                            primaryText: results[result]["Name"],
                            id: results[result]["Id"]
                        });
                    }
                    filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
                    filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
                    resolve(this._convertResultsToPromise(filteredPersonas));
                });
            } else {
                resolve([]);
            }
        });
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
    private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }
}