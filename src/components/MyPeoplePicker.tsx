import * as React from 'react';
import {
  CompactPeoplePicker,
  IContextualMenuItem,
  Dropdown,
  IDropdownOption,
  IPersonaProps,
  IBasePickerSuggestionsProps,
  BaseComponent,
} from '../../node_modules/office-ui-fabric-react/lib/index';
import { assign } from '../../node_modules/office-ui-fabric-react/lib/utilities';
import { IPersonaWithMenu } from '../../node_modules/office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';

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
    //i think the people module needs to be a call to the ups or people picker rest api...
    people.forEach((persona: IPersonaProps) => {
      let target: IPersonaWithMenu = {};

      assign(target, persona, { menuItems: this.contextualMenuItems });
      this._peopleList.push(target);
    });
    this.state = {
      delayResults: false
    };
  }

  public render() {
    return (
      <CompactPeoplePicker
        onResolveSuggestions={ this._onFilterChanged }
        getTextFromItem={ (persona: IPersonaProps) => persona.primaryText }
        pickerSuggestionsProps={ suggestionProps }
        className={ 'ms-PeoplePicker' }
        />
    );
  }

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

  private _toggleChange(toggleState: boolean) {
    this.setState({ delayResults: toggleState });
  }

  private _dropDownSelected(option: IDropdownOption) {
    this.setState({ currentPicker: option.key });
  }

}