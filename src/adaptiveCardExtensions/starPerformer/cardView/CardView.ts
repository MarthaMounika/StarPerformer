import { BaseImageCardView, IImageCardParameters} from '@microsoft/sp-adaptive-card-extension-base';
import { IStarPerformerAdaptiveCardExtensionProps, IStarPerformerAdaptiveCardExtensionState } from '../StarPerformerAdaptiveCardExtension';

export class CardView extends BaseImageCardView<IStarPerformerAdaptiveCardExtensionProps, IStarPerformerAdaptiveCardExtensionState> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */

  public get data(): IImageCardParameters {
    const size: number = this.state.items.length;
    let printText: string;
    if (size > 0){
      printText = "Congratulations!\n" + this.state.items[size-1].title + "\nYou inspire us..";
      if (this.state.items[size-1].title == "NO_ONE"){
        printText = "Sorry, No one is selected as Star Performer";
      }
    } else {
      printText = "Sorry, No one is selected as Star Performer";
    }
    
    return {
      primaryText: printText,
      imageUrl: require('../assets/StarPerformer.jpg'),
      title: this.properties.title
    };
  }
}
