import * as React from 'react';
import { IHelloWorldProps } from './IHelloWorldProps';
interface IHelloWorldState {
    Cars: any;
}
export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
    constructor(props: IHelloWorldProps);
    spServices: any;
    state: {
        Cars: any[];
    };
    renderCars: () => void;
    render(): React.ReactElement<IHelloWorldProps>;
}
export {};
//# sourceMappingURL=HelloWorld.d.ts.map