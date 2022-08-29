import * as React from 'react';
import { IListInfo } from '@pnp/sp/lists';
import { ISearchResult } from '@pnp/sp/search';
import { ISitesService, SitesService } from '../../SitesServices';
import { ServiceScope } from '@microsoft/sp-core-library';

export default class HelloWorld extends React.Component<{ serviceScope: ServiceScope, siteUrl: string }, { sites: ISearchResult[], lists: IListInfo[], error?: string}> {
  private _sitesService: ISitesService;

  
  public constructor(props: { serviceScope: ServiceScope, siteUrl: string}) {
    super(props);

    this._sitesService = props.serviceScope.consume(SitesService.serviceKey);

    this.state = {
      lists: [],
      sites: []
    }
  }

  public componentDidMount(): void {
    this._sitesService.search("*").then((sites) => {
      this.setState({sites});
    }, (error) => {
      this.setState({error});
    });

  }

  public render(): React.ReactElement<{ serviceScope: ServiceScope, siteUrl: string}> {
  
    const { lists, error, sites } = this.state;

    return (
      <section>        
        <div>
          <p>Current time: {(new Date()).toISOString()}</p>
          { error && <p>{JSON.stringify(error)}</p> }
          
          <h3>Sites</h3>
          <ul>
          {
            sites.map(site => <>
              <li>{site.Title} <small>{site.Path}</small></li>
            </>)
          }
          </ul>  
        </div>
      </section>
    );
  }
}
