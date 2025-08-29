<div align="center">

# Redux Example Web Part

The Redux Example Web Part demonstrates using Redux Toolkit for state management in a SharePoint Framework (SPFx) solution.

</div>

<div align="center">

[![SPFx version](https://img.shields.io/badge/SPFx-1.21.1-038387.svg)](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
[![React version](https://img.shields.io/badge/React-17.0.1-087ea4.svg)](https://17.reactjs.org/docs/getting-started.html)
[![Redux Toolkit version](https://img.shields.io/badge/Redux_Toolkit-2.5.0-purple.svg)](https://redux-toolkit.js.org/)
[![Node version](https://img.shields.io/badge/node-22.19.0-026e00.svg)](https://nodejs.org/dist/v22.19.0/)

</div>

<div align="center">

![Redux Example Screenshot](./redux-sample-screenshot.png)
Redux Example Web Part

</div>

## Prerequisites

- Graph Permissions User.Read.All

## Quick Start - Development

- Clone the repository
- CD to the solution directory
- In the command line run
  - **npm install**
  - **npm run serve**
- Navigate to a SharePoint page you would like to test the solution
  - append the following query string to the site _?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js_
  - when prompted select _Load debug files_ button
- Edit the page
  - Add the Redux Example Web Part to the page

## Features

The Redux Example Web Part illustrates the following concepts:

- Using Redux Toolkit for state management in SPFx
- Async state updates with createAsyncThunk
- Listener middleware for side effects
- Integration with Microsoft Graph
- React functional components
- Leveraging MGT React components

## Solution

| Solution Name | Author(s)                 |
| ------------- | ------------------------- |
| Redux Example | Anthony Poulin (ynot3363) |

## Version history

| Version | Date            | Comments        |
| ------- | --------------- | --------------- |
| 1.0     | August 29, 2025 | Initial release |

## Disclaimer

> **THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
