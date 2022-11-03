var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { PrimaryButton } from 'office-ui-fabric-react';
import { SharePointServices } from '../../../Services/SharePointServices';
var HelloWorld = /** @class */ (function (_super) {
    __extends(HelloWorld, _super);
    function HelloWorld(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            Cars: Array(),
        };
        _this.renderCars = function () {
            _this.spServices = new SharePointServices(_this.props.context);
            _this.spServices.getListData("Cars", "$select=Id,Title,Make,Category")
                .then(function (res) {
                res.json().then(function (data) {
                    _this.setState({
                        Cars: data.value,
                    });
                });
            });
        };
        return _this;
    }
    HelloWorld.prototype.render = function () {
        return (React.createElement("section", { className: styles.helloWorld },
            React.createElement("h1", null, "Carss"),
            React.createElement(PrimaryButton, { text: "View" })));
    };
    return HelloWorld;
}(React.Component));
export default HelloWorld;
//# sourceMappingURL=HelloWorld.js.map