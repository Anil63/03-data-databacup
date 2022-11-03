import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { PrimaryButton } from 'office-ui-fabric-react';
import { SharePointServices } from '../../../Services/SharePointServices';


interface IHelloWorldState {
  Cars:any
}

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
 
  constructor(props:IHelloWorldProps){
    super(props);
  }

  spServices :any;
  state = {
    Cars:Array(),
  }

  renderCars = () => {
    this.spServices = new SharePointServices(this.props.context)
      this.spServices.getListData("Cars", "$select=Id,Title,Make,Category")
      .then((res: any) => {
        res.json().then((data: any) => {
          this.setState({
            Cars: data.value,
          });
        });
      });
    };

  public render(): React.ReactElement<IHelloWorldProps> {
    
    return (
      <section className={styles.helloWorld}>
      <h1>Carss</h1>
       <PrimaryButton text="View"   />
      </section>
    );
  }
}
