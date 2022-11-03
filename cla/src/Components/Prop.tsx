import {
  IStackItemProps,
  IStackStyles,
  Label,
  PrimaryButton,
  Stack,
  TextField,
} from "@fluentui/react";
import { type } from "@testing-library/user-event/dist/type";
import React, { Component } from "react";

interface userdata {
  fullname: string;
  emial: string;
  mobile:string;
  password: string;
}
const defaultdata: userdata = {
  fullname: "",
  emial: "",
  mobile:"",
  password: "",
};

const stacktoken = { childrenGap: 15 };
const stackStyles: Partial<IStackStyles> = { root: { with: 800 } };
const columnProps: Partial<IStackItemProps> = {
  styles: { root: { width: 4000 } },
};


export class Prop extends Component  {
  constructor(props: any) {
    super(props);
    this.state = {
      fullname: "",
      emial: "",
      mobile: "",
      password: "",
    };
  }
  defaultdata = {
    fullname:"",
    emial:"",
    mobile:"",
    password:"",
  }
  changeHandler = (event:any) => {
    this.setState({[event.target.id]:event.target.value})
  }
 onsubmithand = (event: React.FormEvent<HTMLElement>) => {
    event.preventDefault();
    console.log(this.state)
    
 
    this.setState({
        fullname:"",
        email:"",
        mobile:"",
        password:"",
    })
  };

  getData = () => {
    localStorage.setItem('userdata', JSON.stringify(this.state))
    const saved:any = localStorage.getItem('userdata');
    const result = JSON.parse(saved)
   this.setState({result})
    
  }


  render() {

    const {fullname,email,mobile,password }:any = this.state
    return (
      <div className="col-md-3 ">
        <h1>User Details</h1>
        <form onSubmit={this.onsubmithand}>
          <div className=" col-md-3 center">
            <Stack
              horizontal
              tokens={stacktoken}
              styles={stackStyles}
              {...columnProps}
            >
              <TextField
                label="FullName"
                placeholder="Full NaMe"
                id="fullname"
                value={fullname} 
                onChange={this.changeHandler}
              />
              <TextField label="Email" placeholder="Email" id="emial" value={email} onChange={this.changeHandler} />
              <TextField label="Mobile" placeholder="Mobile" id="mobile" value={mobile} onChange={this.changeHandler}/>
              <TextField
                label="password"
                type="password"
                canRevealPassword
                revealPasswordAriaLabel="Show password"
                id="password"  
                placeholder="Password"
                value={password}
                onChange={this.changeHandler} />
              <Stack>
                <Label>Add Data</Label>
                <PrimaryButton type="submit"> Insert</PrimaryButton>
              </Stack>
            </Stack>
          </div>
        </form>
      <div className="">
        
      </div>
      
      </div>
    );
  }
}

export default Prop;
