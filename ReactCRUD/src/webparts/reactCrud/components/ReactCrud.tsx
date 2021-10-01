import * as React from 'react';
import styles from './ReactCrud.module.scss';
import { IReactCrudProps } from './IReactCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IReactCrudStates } from './IReactCrudStates';
import {SPOperations} from '../../reactCrud/Services/SPServices';
import { Dropdown, IDropdownOption} from 'office-ui-fabric-react';

export default class ReactCrud extends React.Component<IReactCrudProps, IReactCrudStates, {}> {
  public _spOps:SPOperations;
  public listTitle:string;
constructor(props:IReactCrudProps){
  super(props);
  this.state={
    listTitles:[],
    status:""
  };  
  this._spOps=new SPOperations();
}
componentDidMount(){
   this._spOps.GetAllLists(this.props.spContext).then((result:IDropdownOption[])=>{
     this.setState({listTitles:result});
   });
}

GetListTitle=(evt:any, data:any)=>{
  this.listTitle=data.text;

}
  public render(): React.ReactElement<IReactCrudProps> {
    return (
         <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts CRUD Demo.</p>            
            </div>
            <div className={styles.myStyles}>
              <Dropdown  className={styles.dropdown} options={this.state.listTitles} placeholder="Select a List" onChange={this.GetListTitle}></Dropdown>
              <button className={styles.button} onClick={()=>this._spOps.CreateListItem(this.props.spContext,this.listTitle).then((result:string)=>{
                this.setState({status:result})
              })}>Create List Item</button>
              <button className={styles.button } onClick={()=>this._spOps.UpdateListItem(this.props.spContext,this.listTitle).then((result:string)=>{
                this.setState({status:result})
              })}>Update List Item</button>
              <button className={styles.button} onClick={()=>this._spOps.DeleteListItem(this.props.spContext,this.listTitle).then((result:string)=>{
                this.setState({status:result})
              })}>Delete List Item</button>
            </div>
           <div>{this.state.status}</div>
          </div>
        </div>
      
    );
  }
}
