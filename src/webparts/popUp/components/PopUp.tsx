import * as React from 'react';
//import styles from './PopUp.module.scss';
import type { IPopUpProps } from './IPopUpProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { IStackItemStyles, IStackStyles, IStackTokens, Stack } from '@fluentui/react/lib/Stack';

import { DefaultEffects, IButtonStyles, Modal, PrimaryButton,  } from '@fluentui/react';
import { IPopUpState } from './IPopUpState';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";


// Styles definition
const textStyle:React.CSSProperties={
  color:'black',
  width:'100%'
}
const stackStyles: IStackStyles = {
  root: {
    boxShadow:DefaultEffects.elevation8
  },
};
const stackItemStyles: IStackItemStyles = {
  root: {
    alignItems: 'center',
    display: 'flex',
    justifyContent: 'center',
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
      buttonText,
      popUpText,
      displayMode
    } = this.props;

    return (
      <>
      <Stack enableScopedSelectors tokens={outerStackTokens}>
        <Stack enableScopedSelectors styles={stackStyles} tokens={innerStackTokens}>
          {displayMode==2 && (
            <Stack.Item grow styles={stackItemStyles}>
              <RichText style={textStyle} value={popUpText} onChange={(text)=>this.props.onTextChange(text)} isEditMode={displayMode===2?true:false} label='PopUp Content'/>
            </Stack.Item>
          )}
          <Stack.Item styles={stackItemStyles}>
            <PrimaryButton text={buttonText} onClick={this.showModal} styles={buttonStyles}/>
          </Stack.Item>
        </Stack>
      </Stack>

      <Modal
        titleAriaId={'1'}
        isOpen={this.state.isModalOpen}
        onDismiss={this.hideModal}
      >
        <RichText style={textStyle} value={popUpText} isEditMode={false} />
      </Modal>
      </>
    );
  }
}
