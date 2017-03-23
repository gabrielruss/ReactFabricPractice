/* tslint:disable */
import * as React from 'react';
import { Promise } from "es6-promise";
/* tslint:enable */
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
    NormalPeoplePicker
} from '../../node_modules/office-ui-fabric-react/lib/Pickers';
import { IPersonaWithMenu } from '../../node_modules/office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { RestUtil } from '../../utils/RestUtil';

//keeping this in for example state
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

    constructor() {
        super();
        this._peopleList = [];
        // people.forEach((persona: IPersonaProps) => {
        //     let target: IPersonaWithMenu = {};

        //     assign(target, persona, { menuItems: this.contextualMenuItems });
        //     this._peopleList.push(target);
        // });
        this.state = {
            delayResults: false
        };
    }

    public render() {
        return (
            <div>
                <CompactPeoplePicker
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

            RestUtil.getUsers(filterText).then((results) => {
            let filteredPersonas:IPersonaProps[] = [];
            //need to have REST call return an array of users
            // call function that will build array of user info based on results
            for (let result in results) {
                filteredPersonas.push({
                    key: result,
                    primaryText: results[result]["Name"]
                });
                console.log(`Found ${results[result]["Name"]}`);
            }
            return filteredPersonas;
        });

            //filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
            //filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
            //below will need to happen to limit results
            //return this._convertResultsToPromise(filteredPersonas);
            
        } else {
            return [];
        }
    }

    // private _filterPersonasByText(filterText: string): IPersonaProps[] {
        
        
    //     //return filteredPersonas;
    // }

    // private _buildUserArray(results: Object, userArray?: IPersonaProps[]) {
        
    // }

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
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }
}