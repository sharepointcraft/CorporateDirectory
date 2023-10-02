import * as React from 'react';
import { IRefinersCardProps } from './IRefinersCardProps';
import { IRefinersCardState } from './IRefinersCardState';
import {
  Log, Environment, EnvironmentType,
} from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  Persona,
  PersonaSize,
  DocumentCard,
  DocumentCardType,
  Icon,
} from 'office-ui-fabric-react';

const EXP_SOURCE: string = 'SPFxDirectory';
const LIVE_PERSONA_COMPONENT_ID: string =
  '914330ee-2df2-4f6e-a858-30c23a812408';

export class RefinersCard extends React.Component<
  IRefinersCardProps,
  IRefinersCardState
  > {
  constructor(props: IRefinersCardProps) {
    super(props);
    console.log(props);
    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }
  /**
   *
   *
   * @memberof PersonaCard
   */
  public async componentDidMount() {
    if (Environment.type !== EnvironmentType.Local) {
      const sharedLibrary = await this._loadSPComponentById(
        LIVE_PERSONA_COMPONENT_ID
      );
      const livePersonaCard: any = sharedLibrary.LivePersonaCard;
      this.setState({ livePersonaCard: livePersonaCard });
    }
  }

  /**
   *
   *
   * @param {IPersonaCardProps} prevProps
   * @param {IPersonaCardState} prevState
   * @memberof PersonaCard
   */
  public componentDidUpdate(
    prevProps: IRefinersCardProps,
    prevState: IRefinersCardState
  ): void { }

  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */
  // private _LivePersonaCard() {
  //   return React.createElement(
  //     this.state.livePersonaCard,
  //     {
  //       serviceScope: this.props.context.serviceScope,
  //       upn: this.props.profileProperties.Email,
  //       onCardOpen: () => {
  //         console.log('LivePersonaCard Open');
  //       },
  //       onCardClose: () => {
  //         console.log('LivePersonaCard Close');
  //       },
  //     },
  //     // this._PersonaCard()
  //   );
  // }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  // private _PersonaCard(): JSX.Element {
  //   return (
  //     <DocumentCard
  //       className={styles.documentCard}
  //       type={DocumentCardType.normal}
  //     >
  //       <div className={styles.persona}>
  //         <Persona
  //           text={this.props.profileProperties.DisplayName}
  //           secondaryText={this.props.profileProperties.Title}
  //           tertiaryText={this.props.profileProperties.Department}
  //           imageUrl={this.props.profileProperties.PictureUrl}
  //           size={PersonaSize.size72}
  //           imageShouldFadeIn={false}
  //           imageShouldStartVisible={true}
  //         >
  //           {this.props.profileProperties.Department ? (
  //             <div className={styles.customClass}>
  //               <Icon iconName="Phone" style={{ fontSize: '12px' }} />
  //               <span style={{ marginLeft: 5, fontSize: '12px' }}>
  //                 {' '}
  //                 {this.props.profileProperties.Department}
  //               </span>
  //             </div>
  //           ) : (
  //               ''
  //             )}
  //           {this.props.profileProperties.WorkPhone ? (
  //             <div>
  //               <Icon iconName="Phone" style={{ fontSize: '12px' }} />
  //               <span style={{ marginLeft: 5, fontSize: '12px' }}>
  //                 {' '}
  //                 {this.props.profileProperties.WorkPhone}
  //               </span>
  //             </div>
  //           ) : (
  //               ''
  //             )}
  //           {this.props.profileProperties.Location ? (
  //             <div className={styles.textOverflow}>
  //               <Icon iconName="Poi" style={{ fontSize: '12px' }} />
  //               <span style={{ marginLeft: 5, fontSize: '12px' }}>
  //                 {' '}
  //                 {this.props.profileProperties.Location}
  //               </span>
  //             </div>
  //           ) : (
  //               ''
  //             )}
  //         </Persona>
  //       </div>
  //     </DocumentCard>
  //   );
  // }
  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      Promise.reject(error);
      Log.error(EXP_SOURCE, error, this.props.context.serviceScope);
    }
  }

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IRefinersCardProps> {
    return (
      <div>
        Hello Refiners
      </div>
    );
  }
}
