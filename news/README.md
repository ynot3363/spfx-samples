# news

## Summary

Demonstrates how to add in additional commands to manage News Posts within a Site's Site Pages library. This extension will provide the abilitiy to Demote News Posts, Promote Pages to News Posts and update the Publish Date of News Posts.

**Demote News Post**
![screenshot of demote news post command]()

**Promote / Update Publishing Date News Post**
![screenshot of promote / update publishing date news post commands]()

**Update Publishing Date Panel**
![screenshot of the panel when updating publising date]()

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.14-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

The solution depends on having a Site Pages library created on the site and the Promoted State column added all the views.

## Solution

| Solution | Author(s)                                                                                                   |
| -------- | ----------------------------------------------------------------------------------------------------------- |
| news     | [Anthony Poulin](https://anthonyepoulin.com) follow me on Twitter [@ynot3363](https://twitter.com/ynot3363) |

## Version history

| Version | Date              | Comments        |
| ------- | ----------------- | --------------- |
| 1.0     | December 31, 2022 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Update the properties in the following files to match your env.
  - config > serve.json
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

The **news** command set will add the ability for users to demote news posts, promote pages to news posts and update the publishing date of news posts (which control the order they appear in the News web part). A user can perform the actions on 1 or many pages. If more than 1 page is selected then all the pages either need to be a news post or a page for the commands to appear.

This extension illustrates the following concepts:

- Demote News Posts to Pages
- Promote Pages to News Posts
- Update the Publishing Date for News Posts
- Leverage React components within a CommandSet
