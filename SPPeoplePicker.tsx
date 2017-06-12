import * as React from 'react';
import { CompactPeoplePicker, IPersonaProps, IBasePickerSuggestionsProps } from 'office-ui-fabric-react';

export interface ISPPeoplePickerProps {
    defaultValues?: IPersonaProps[];
    multi?: boolean;
    onChange?(people: IPersonaProps[]): void;
}

export default class SPPeoplePicker extends React.Component<ISPPeoplePickerProps, any> {
    constructor() {
        super();
        this._onFilterChanged = this._onFilterChanged.bind(this); // https://github.com/goatslacker/alt/issues/283
        this._onStateChange = this._onStateChange.bind(this);
        this.state = {};
        if (typeof (_spPageContextInfo) !== 'undefined') {
            this.state.pageContext = _spPageContextInfo;
        } else {
            this.state.pageContext = {
                siteServerRelativeUrl: '',
                serverRequestPath: window.location.href
            };
        }
    }

    public componentDidMount(): void {
        this.state.pickerEnabled = (!(this.props.defaultValues != null && this.props.defaultValues.length > 0)) || this.props.multi;
    }

    public render(): React.ReactElement<null> {
        const suggestionProps: IBasePickerSuggestionsProps = {
            noResultsFoundText: 'No results found',
            loadingText: 'Loading'
        };

        const pickerDisplayProps: React.HTMLProps<HTMLInputElement> = {
            disabled: !this.state.pickerEnabled
        };
        return (
            <div>
                <CompactPeoplePicker onResolveSuggestions={this._onFilterChanged} onChange={this._onStateChange} className={'ms-PeoplePicker'} pickerSuggestionsProps={suggestionProps} inputProps={pickerDisplayProps} defaultSelectedItems={this.props.defaultValues} />
                <p />
            </div>
        );
    }

    private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        if (filterText) {
            return this._getResultsAsPromise(filterText);
        } else {
            return [];
        }
    }

    private _onStateChange(currentPersonas: IPersonaProps[]) {
        this.setState({ pickerEnabled: currentPersonas.length === 0 || this.props.multi });
        if (this.props.onChange) {
            this.props.onChange(currentPersonas);
        }
    }

    private _getResultsAsPromise(filterText: string): Promise<IPersonaProps[]> {
        return new Promise<IPersonaProps[]>((resolve, reject) => {
            let url = `${this.state.pageContext.siteServerRelativeUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;

            var query = { 'queryParams': { 'QueryString': filterText, 'MaximumEntitySuggestions': 50, 'AllowEmailAddresses': false, 'AllowOnlyEmailAddresses': false, 'PrincipalType': 1, 'PrincipalSource': 1, 'SharePointGroupID': 0 } };

            if (window.location.host.indexOf('localhost') !== -1) {
                return resolve([{ primaryText: 'Awesome pants', key: 'account@account' }]);
            }

            fetch(url, {
                method: 'POST',
                headers: {
                    'Accept': 'application/json;odata=minimalmetadata',
                    'Content-Type': 'application/json;odata=minimalmetadata',
                    'Cache': 'no-cache',
                    'X-RequestDigest': this.state.pageContext.formDigestValue
                },
                credentials: 'include',
                body: JSON.stringify(query)
            }).then((res) => {
                return res.json();
            }).then((suggestions: any) => {
                let people: any[] = JSON.parse(suggestions.value);
                let personas: IPersonaProps[] = [];

                for (var i = 0; i < people.length; i++) {
                    var p = people[i];
                    let s: IPersonaProps = {};
                    let account = p.Key.substr(p.Key.lastIndexOf('|') + 1);
                    s.primaryText = p.DisplayText;
                    s.key = p.Key;
                    s.imageUrl = `/_layouts/15/userphoto.aspx?size=S&accountname=${account}`;
                    s.imageShouldFadeIn = true;
                    personas.push(s);
                }
                return resolve(personas);
            }).catch(() => {
                return reject([]);
            });
        });
    }
}