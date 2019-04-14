# **Developing Office 365 Outlook Add-in with [Office UI Fabric React](https://developer.microsoft.com/en-us/fabric#/get-started#react)**

## **Note:**
- Outlook add-in works only for Office365 users and not for any other users with Gmail or simple outlook.com accounts.
- Outlook add-in can be developded using one of the pure JS, React JS, Angular JS or Vue JS
- Before developing any add-on for production environment, please go through best practices suggested by Microsoft  
https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/add-in-development-best-practices   
https://docs.microsoft.com/en-us/office/dev/add-ins/design/ux-design-pattern-templates
 
## **Technology Stack**
1) React - Office UI Fabric React - For the UI design

## How to make a boiler-plate template?
There are two ways to make the boiler plat template for this kind of project.

1) Using create-react-app and
2) Using Yeoman project generator

We are going to build our project using create-react-app

## **Prerequisites**
1)	Install create-react-app
```
npm install -g create-react-app
```
2)	Install Yeoman
```
npm install -g yo
```

3) Install Office Addin  Project Creator

```
npm install -g yo generator-office
```

## Creating the outlook Add-in App and Installing the Dependencies 

Run following commands on the Command Prompt

1) Generate the new React App

```
 npx create-react-app outlook_leave_addin_app
```
2) Navigate to the outlook_leave_addin_app folder

```
cd outlook_leave_addin_app
```

3) Follow the steps [here](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react#generate-the-manifest-file-and-sideload-the-add-in) (**select outlook for "Which Office client application would you like to support" question**) to create the office add-in manifest file.

4) Install "office-js".

```
npm install @micrsoft\office-js --save   
```

Ref:(https://github.com/OfficeDev/office-js)

5) Install the Fabric React package 

```
npm --save install office-ui-fabric-react
```

6) For styling run following command to install required packages

```
npm install --save @uifabric/styling
```

7) Run the following command to use the [Office fabric UI Layouts](https://developer.microsoft.com/en-us/fabric#/styles/layout)

```
npm install office-ui-fabric-core
```
8) Run following command if you want to use axios.

```
npm install axios
```

Note: To run the React app locally with HTTPS, type the following command in the location of the root folder.Because, add-ins communicate only with https end points. 


```
set HTTPS=true&&npm start
```

## How to Build the Addin 

1) Put following setting in the ```package.json```. This setting will make sure to refer the "js" and "css" files in the ```assets``` folder using the relative path in the index.html of the build package. [More details are here.](https://github.com/facebook/create-react-app/blob/master/packages/react-scripts/template/README.md#serving-the-same-build-from-different-paths)

```
put this in your package.json
```

2) Run the following command to build the package

```
npm run build
```

3) Build package will created be in the ```build``` folder.

### References

https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins?context=outlook/context

https://docs.microsoft.com/en-us/outlook/

https://docs.microsoft.com/en-us/outlook/add-ins/

https://docs.microsoft.com/en-us/outlook/add-ins/quick-start?tabs=visual-studio-code

https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial

https://docs.microsoft.com/en-us/office/dev/add-ins/concepts/add-in-development-lifecycle

https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox.item

https://github.com/OfficeDev/office-js-docs/blob/master/reference/

https://github.com/OfficeDev/office-ui-fabric-react/tree/master/packages/office-ui-fabric-react/src/components

https://docs.microsoft.com/en-us/office/dev/add-ins/develop/use-the-oauth-authorization-framework-in-an-office-add-in

https://developer.microsoft.com/en-us/office/docs

https://github.com/OfficeDev/office-ui-fabric-react/blob/master/packages/icons/README.md

https://github.com/OfficeDev/office-ui-fabric-react/tree/master/packages/styling

https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-get-started-react

https://github.com/OfficeDev/office-js

https://developer.microsoft.com/en-us/fabric#/get-started
