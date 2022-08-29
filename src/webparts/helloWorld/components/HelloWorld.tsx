import * as React from 'react';
import { IHelloWorldProps } from './IHelloWorldProps';
import { IListInfo } from '@pnp/sp/lists';
import { ISearchResult } from '@pnp/sp/search';
import { ISitesService, SitesService } from '../../../SitesServices';

export default class HelloWorld extends React.Component<IHelloWorldProps, { sites: ISearchResult[], lists: IListInfo[], error?: string}> {
  private _sitesService: ISitesService;

  
  public constructor(props: IHelloWorldProps) {
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

  public render(): React.ReactElement<IHelloWorldProps> {   
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
