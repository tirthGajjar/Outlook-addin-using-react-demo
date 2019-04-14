/* global Office */
import * as React from 'react';
// import './App.css';

//Importing Layout classes  https://developer.microsoft.com/en-us/fabric#/styles/layout
// import 'office-ui-fabric-core/dist/css/fabric.min.css'

import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { initializeIcons } from '@uifabric/icons';

class App extends React.Component {
    officeMailBoxItem = Office.context.mailbox.item;

    constructor(props) {
        super(props);
        this.setSubject("Tracked Email");

        this.getEmailList();

        //Without this line Datepicker calendar icon and dropdown caret icon does not display
        initializeIcons();
    }

    emailDetails = {
        subject: '',
        reason: '',
        startDate: new Date(),
        endDate: new Date(),
        leaveType: {
            text: ''
        },
    };

    setSubject = (subject) => {
        this.emailDetails.subject = subject;
        this.officeMailBoxItem.subject.setAsync(subject);
    };

    setStartDate = (value) => {
        let date = new Date(value);
        this.emailDetails.startDate = date;
    };

    setEndDate = (value) => {
        let date = new Date(value);
        this.emailDetails.endDate = date;
    };

    setLeaveType = (option) => {
        this.emailDetails.leaveType = option;
    };

    setReason = (value) => {
        this.emailDetails.reason = value;
    };

    createMessage = () => {
        this.officeMailBoxItem.body.prependAsync(
            '<p>' +
            ' Hi, <br/>' +
            'I am on ' + this.emailDetails.leaveType.text + ' from ' + this.emailDetails.startDate.toLocaleDateString() + ' to ' + this.emailDetails.endDate.toLocaleDateString() +
            ', because ' + this.emailDetails.reason + '. <br/> <br/>' +
            'Thank you.' +
            '</p>',
            { coercionType: Office.CoercionType.Html })
    };

    getEmailList = () => {

        let setEmail = () => {
            this.officeMailBoxItem.to.setAsync(["tirthg@saleshandyonmicrosoft.com"]);
        };
        setEmail();
    };

    render() {
        return (
            <div id="addInContainerDiv" className='ms-grid'>

                <div className="ms-Grid-row" >
                    <DatePicker className="ms-Grid-col ms-sm12 ms-lg4"
                        placeholder='Select the From Date'
                        onSelectDate={
                            this.setStartDate
                        }
                    />
                </div>

                <div className="ms-Grid-row" >
                    <DatePicker className="ms-Grid-col ms-sm12 ms-lg4"
                        placeholder='Select the To Date'
                        onSelectDate={
                            this.setEndDate
                        }
                    />
                </div>

                <div className="ms-Grid-row" >
                    <Dropdown className="ms-Grid-col ms-sm12 ms-lg4"
                        placeHolder='Select Leave Type'
                        onChanged={
                            this.setLeaveType
                        }
                        options={
                            [
                                { key: 'annual', text: 'Annual Leave' },
                                { key: 'personal', text: 'Sick/Carer Leave' }
                            ]} > </Dropdown>
                </div>

                <div className="ms-Grid-row" >
                    <TextField className="ms-Grid-col ms-sm12 ms-lg4"
                        label='Reason'
                        onChanged={
                            this.setReason
                        }
                        multiline
                        rows={5}
                        autoAdjustHeight
                    />
                </div>
                <div className="ms-Grid-row" >
                    <div id="buttonContainerDiv" className="ms-Grid-col ms-sm12 ms-lg4">
                        <DefaultButton id='OKButton' primary={true} onClick={this.createMessage} >
                            Apply
              </DefaultButton>
                    </div>
                </div>
            </div>
        );
    }
}

export default App;