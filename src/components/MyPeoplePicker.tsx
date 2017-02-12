/* tslint:disable */
import * as React from 'react';
import { Promise } from "es6-promise";
/* tslint:enable */
import {
    BaseComponent,
    assign,
    autobind
} from '../../node_modules/office-ui-fabric-react/lib/Utilities';
import { Dropdown, IDropdownOption } from '../../node_modules/office-ui-fabric-react/lib/index';
import { Toggle } from '../../node_modules/office-ui-fabric-react/lib/index';
import { IPersonaProps } from '../../node_modules/office-ui-fabric-react/lib/Persona';
import { IContextualMenuItem } from '../../node_modules/office-ui-fabric-react/lib/ContextualMenu';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    ListPeoplePicker,
    NormalPeoplePicker
} from '../../node_modules/office-ui-fabric-react/lib/Pickers';
import { IPersonaWithMenu } from '../../node_modules/office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { people } from '../../node_modules/office-ui-fabric-react/lib/demo/pages/PeoplePickerPage/examples/PeoplePickerExampleData';
import '../../node_modules/office-ui-fabric-react/lib/demo/pages/PeoplePickerPage/examples/PeoplePicker.Types.Example.scss';

export interface IPeoplePickerExampleState {
    currentPicker?: number | string;
    delayResults?: boolean;
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};

export class PeoplePickerTypesExample extends BaseComponent<any, IPeoplePickerExampleState> {
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
        people.forEach((persona: IPersonaProps) => {
            let target: IPersonaWithMenu = {};

            assign(target, persona, { menuItems: this.contextualMenuItems });
            this._peopleList.push(target);
        });
        this.state = {
            currentPicker: 1,
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

    public _renderPreselectedItemsPicker() {
        return (
            <CompactPeoplePicker
                onResolveSuggestions={this._onFilterChanged}
                getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
                className={'ms-PeoplePicker'}
                defaultSelectedItems={people.splice(0, 3)}
                key={'list'}
                pickerSuggestionsProps={suggestionProps}
            />
        );
    }

    public _renderLimitedSearch() {
        let limitedSearchSuggestionProps = suggestionProps;
        limitedSearchSuggestionProps.searchForMoreText = 'Load all Results';
        return (
            <CompactPeoplePicker
                onResolveSuggestions={this._onFilterChangedWithLimit}
                getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
                className={'ms-PeoplePicker'}
                onGetMoreResults={this._onFilterChanged}
                pickerSuggestionsProps={limitedSearchSuggestionProps}
            />
        );
    }

    @autobind
    private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        if (filterText) {
            let filteredPersonas: IPersonaProps[] = this._filterPersonasByText(filterText);

            filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
            filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
            return this._filterPromise(filteredPersonas);
        } else {
            return [];
        }
    }

    @autobind
    private _onFilterChangedWithLimit(filterText: string, currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        return this._onFilterChanged(filterText, currentPersonas, 3);
    }

    private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        if (this.state.delayResults) {
            return this._convertResultsToPromise(personasToReturn);
        } else {
            return personasToReturn;
        }
    }

    private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.primaryText === persona.primaryText).length > 0;
    }

    private _filterPersonasByText(filterText: string): IPersonaProps[] {
        return this._peopleList.filter(item => this._doesTextStartWith(item.primaryText, filterText));
    }

    private _doesTextStartWith(text: string, filterText: string): boolean {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    }

    private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }

    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }

    @autobind
    private _toggleChange(toggleState: boolean) {
        this.setState({ delayResults: toggleState });
    }

    @autobind
    private _dropDownSelected(option: IDropdownOption) {
        this.setState({ currentPicker: option.key });
    }

}