import * as React from 'react';
//import styles from './PopUp.module.scss';
import type { IPopUpProps } from './IPopUpProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { IStackItemStyles, IStackStyles, IStackTokens, Stack } from '@fluentui/react/lib/Stack';
import { DefaultPalette } from '@fluentui/react/lib/Theme';
import { IButtonStyles, Modal, PrimaryButton } from '@fluentui/react';
import { IPopUpState } from './IPopUpState';



// Styles definition
const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themePrimary,
    //height: 250,
  },
};
const stackItemStyles: IStackItemStyles = {
  root: {
    alignItems: 'center',
    background: DefaultPalette.themeLighterAlt,
    color: DefaultPalette.white,
    display: 'flex',
    justifyContent: 'center',
    margin:15
  },
};
const buttonStyles:IButtonStyles = {
  root:{
    margin:15
  }
}
// Tokens definition
const outerStackTokens: IStackTokens = { childrenGap: 5 };
const innerStackTokens: IStackTokens = {
  childrenGap: 5,
  padding: 10,
};
export default class PopUp extends React.Component<IPopUpProps, IPopUpState> {

  constructor(props:IPopUpProps){

    super(props);

    this.state={
      isModalOpen:false
    }
    this.hideModal = this.hideModal.bind(this);
    this.showModal = this.showModal.bind(this);

  }
  private hideModal():void {
    this.setState({
      isModalOpen:false
    })
  }
  private showModal():void {
    this.setState({
      isModalOpen:true
    })
  }

  public render(): React.ReactElement<IPopUpProps> {
    const {
      //description,
      //isDarkTheme,
      //environmentMessage,
      //hasTeamsContext,
      //userDisplayName
    } = this.props;

    return (
      <>
      <Stack enableScopedSelectors tokens={outerStackTokens}>
        <Stack enableScopedSelectors styles={stackStyles} tokens={innerStackTokens}>
          <Stack.Item grow={3} styles={stackItemStyles}>
            <PrimaryButton text='text' onClick={this.showModal} styles={buttonStyles}/>
          </Stack.Item>
        </Stack>
      </Stack>

      <Modal
        titleAriaId={'1'}
        isOpen={this.state.isModalOpen}
        onDismiss={this.hideModal}
        //containerClassName={contentStyles.container}
        //dragOptions={isDraggable ? dragOptions : undefined}
      >
        
      </Modal>
      </>
    );
  }
}
