import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { StarPerformerPropertyPane } from './StarPerformerPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http-base';

export interface IStarPerformerAdaptiveCardExtensionProps {
  title: string;
  listId: string;
}

export interface IStarPerformerAdaptiveCardExtensionState {
  items: IListItem[];
}

export interface IListItem {
  title: string;
  email: string;
  description: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'StarPerformer_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'StarPerformer_QUICK_VIEW';

export default class StarPerformerAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IStarPerformerAdaptiveCardExtensionProps,
  IStarPerformerAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: StarPerformerPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      items: []
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'StarPerformer-property-pane'*/
      './StarPerformerPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.StarPerformerPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  private _fetchData(): Promise<void> {
    if (this.properties.listId) {
      const mainSiteUrl: string = this.context.pageContext.site.absoluteUrl.replace(this.context.pageContext.site.serverRelativeUrl, '');
      const apiUrl: string = `${mainSiteUrl}/_api/web/lists/GetById('${this.properties.listId}')/items`;
      console.log(apiUrl);
      return this.context.spHttpClient.get(
        apiUrl, SPHttpClient.configurations.v1
      )
        .then((response) => response.json())
        .then((jsonResponse) => jsonResponse.value.map(
          (item: { Title: any; Description: any; Description0: any}) => { return { 
            title: item.Title, email: item.Description, description: item.Description0}; })
          )
        .then((items) => this.setState({ items }));
    }
  
    return Promise.resolve();
  }
}
