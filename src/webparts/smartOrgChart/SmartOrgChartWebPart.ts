import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';

import { SmartOrgChart } from './components/SmartOrgChart';
import { ISmartOrgChartProps, OrgChartTheme } from './components/ISmartOrgChartProps';

export interface ISmartOrgChartWebPartProps {
  defaultView: 'directory' | 'orgchart';
  topLevelUser: string;
  levelsBelow: number;
  pageSize: number;
  useDemoData: boolean;
  // Branding
  companyName: string;
  logoUrl: string;
  // Visual style (admin-controlled)
  theme: OrgChartTheme;
  defaultLayout: 'drill' | 'vertical' | 'horizontal';
  // Feature flags
  enableFindMe: boolean;
  enableLayoutToggle: boolean;
  enableStats: boolean;
  enableDeptFilter: boolean;
  enableUserFilter: boolean;
  // Data
  dataSource: 'auto' | 'graph' | 'search';
  // User filters
  excludedAccounts: string;
  restrictToTenantDomain: boolean;
  hideGuestUsers: boolean;
  hideDisabledAccounts: boolean;
}

export default class SmartOrgChartWebPart extends BaseClientSideWebPart<ISmartOrgChartWebPartProps> {
  protected onInit(): Promise<void> {
    const p = this.properties;
    // Feature flag defaults
    if (p.enableFindMe       === undefined) p.enableFindMe       = true;
    if (p.enableLayoutToggle === undefined) p.enableLayoutToggle = true;
    if (p.enableStats        === undefined) p.enableStats        = true;
    if (p.enableDeptFilter   === undefined) p.enableDeptFilter   = true;
    if (p.enableUserFilter   === undefined) p.enableUserFilter   = true;
    // Visual style defaults
    if (p.theme         === undefined) p.theme         = 'modern';
    if (p.defaultLayout === undefined) p.defaultLayout = 'drill';
    // Branding defaults
    if (p.companyName === undefined) p.companyName = '';
    if (p.logoUrl     === undefined) p.logoUrl     = '';
    // Data source default
    if (p.dataSource  === undefined) p.dataSource  = 'auto';
    // User filter defaults
    if (p.excludedAccounts       === undefined) p.excludedAccounts       = '';
    if (p.restrictToTenantDomain === undefined) p.restrictToTenantDomain = false;
    if (p.hideGuestUsers         === undefined) p.hideGuestUsers         = true;
    if (p.hideDisabledAccounts   === undefined) p.hideDisabledAccounts   = true;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ISmartOrgChartProps> = React.createElement(SmartOrgChart, {
      context: this.context,
      defaultView: this.properties.defaultView || 'directory',
      topLevelUser: this.properties.topLevelUser || '',
      levelsBelow: this.properties.levelsBelow || 3,
      pageSize: this.properties.pageSize || 50,
      useDemoData: this.properties.useDemoData || false,
      companyName: this.properties.companyName || '',
      logoUrl: this.properties.logoUrl || '',
      theme: this.properties.theme || 'modern',
      defaultLayout: this.properties.defaultLayout || 'drill',
      enableFindMe: this.properties.enableFindMe !== false,
      enableLayoutToggle: this.properties.enableLayoutToggle !== false,
      enableStats: this.properties.enableStats !== false,
      enableDeptFilter: this.properties.enableDeptFilter !== false,
      enableUserFilter: this.properties.enableUserFilter !== false,
      dataSource: this.properties.dataSource || 'auto',
      excludedAccounts:       this.properties.excludedAccounts       || '',
      restrictToTenantDomain: this.properties.restrictToTenantDomain || false,
      hideGuestUsers:         this.properties.hideGuestUsers         || false,
      hideDisabledAccounts:   this.properties.hideDisabledAccounts   || false,
      onSettingsSaved: (newProps: Partial<ISmartOrgChartWebPartProps>) => {
        Object.assign(this.properties, newProps);
      }
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Smart Org Chart Settings' },
          groups: [
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneToggle('useDemoData', {
                  label: 'Use Demo Data',
                  onText: 'On — showing 125 sample employees',
                  offText: 'Off — live Microsoft 365 data',
                  checked: false
                }),
                PropertyPaneChoiceGroup('defaultView', {
                  label: 'Default View',
                  options: [
                    { key: 'directory', text: 'Employee Directory', iconProps: { officeFabricIconFontName: 'People' } },
                    { key: 'orgchart', text: 'Org Chart', iconProps: { officeFabricIconFontName: 'Org' } }
                  ]
                })
              ]
            },
            {
              groupName: 'Branding',
              groupFields: [
                PropertyPaneTextField('companyName', {
                  label: 'App Title',
                  placeholder: 'Contoso Org Chart',
                  description: 'Displayed in the header bar alongside the view name'
                }),
                PropertyPaneTextField('logoUrl', {
                  label: 'Logo URL',
                  placeholder: 'https://contoso.sharepoint.com/sites/yoursite/SiteAssets/logo.png',
                  description: 'Full URL to a PNG/SVG/JPG image. Tip: open the image file in your browser and copy the address bar URL. "Copy link" sharing URLs will not work.'
                }),
              ]
            },
            {
              groupName: 'Visual Style',
              groupFields: [
                PropertyPaneChoiceGroup('theme', {
                  label: 'Chart Theme',
                  options: [
                    { key: 'modern',    text: 'Modern — dept colors on white',  iconProps: { officeFabricIconFontName: 'Color' } },
                    { key: 'minimal',   text: 'Minimal — flat & low contrast',  iconProps: { officeFabricIconFontName: 'CollapseMenu' } },
                    { key: 'corporate', text: 'Corporate — unified blue',        iconProps: { officeFabricIconFontName: 'Work' } },
                    { key: 'dark',      text: 'Dark — dark navy background',     iconProps: { officeFabricIconFontName: 'ClearNight' } },
                  ]
                }),
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: 'Default Org Chart Layout',
                  options: [
                    { key: 'drill',      text: 'Drill-Down (recommended)',  iconProps: { officeFabricIconFontName: 'Org' } },
                    { key: 'vertical',   text: 'Vertical Tree',             iconProps: { officeFabricIconFontName: 'ArrowUpRight' } },
                    { key: 'horizontal', text: 'Horizontal Tree',           iconProps: { officeFabricIconFontName: 'ArrowDownRight' } },
                  ]
                }),
              ]
            },
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneChoiceGroup('dataSource', {
                  label: 'Where to load user & org data from',
                  options: [
                    {
                      key: 'auto',
                      text: 'Auto — Graph API, fall back to SharePoint Search',
                      iconProps: { officeFabricIconFontName: 'AutoEnhanceOn' }
                    },
                    {
                      key: 'graph',
                      text: 'Graph API — live Azure AD data (no indexing delay)',
                      iconProps: { officeFabricIconFontName: 'AzureLogo' }
                    },
                    {
                      key: 'search',
                      text: 'SharePoint Search — legacy behavior',
                      iconProps: { officeFabricIconFontName: 'Search' }
                    }
                  ]
                }),
                PropertyPaneLabel('dataSource', {
                  text: 'Graph API is recommended. It reads directly from Azure Active Directory so new users and manager changes appear immediately. Requires Microsoft Graph permissions to be approved in the SharePoint App Catalog.'
                })
              ]
            },
            {
              groupName: 'User Filters',
              groupFields: [
                PropertyPaneTextField('excludedAccounts', {
                  label: 'Exclude accounts',
                  placeholder: 'conf-room, noreply, admin@, Service Account',
                  description: 'Comma-separated words or patterns (case-insensitive). Any user whose display name, email, or UPN contains one of these will be hidden from all views.',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneToggle('restrictToTenantDomain', {
                  label: 'Only show tenant users',
                  onText: 'On — hides accounts with external email domains (e.g. gmail.com, hotmail.com)',
                  offText: 'Off — all users shown regardless of email domain'
                }),
                PropertyPaneToggle('hideGuestUsers', {
                  label: 'Hide Azure AD guest accounts',
                  onText: 'On — guest accounts hidden',
                  offText: 'Off — guest accounts visible (shown with Guest badge)'
                }),
                PropertyPaneToggle('hideDisabledAccounts', {
                  label: 'Hide disabled accounts',
                  onText: 'On — blocked sign-in accounts hidden',
                  offText: 'Off — disabled accounts visible (shown with Disabled badge)'
                }),
              ]
            },
            {
              groupName: 'Org Chart',
              groupFields: [
                PropertyPaneTextField('topLevelUser', {
                  label: 'Top-Level User (UPN or Email)',
                  placeholder: 'ceo@yourcompany.com',
                  description: 'The person shown at the root of the org chart'
                }),
                PropertyPaneSlider('levelsBelow', {
                  label: 'Levels to load below root',
                  min: 1,
                  max: 8,
                  value: 3,
                  showValue: true,
                  step: 1
                })
              ]
            },
            {
              groupName: 'Org Chart Features',
              groupFields: [
                PropertyPaneLabel('enableFindMe', {
                  text: 'Show or hide toolbar buttons in the org chart view.'
                }),
                PropertyPaneToggle('enableFindMe', {
                  label: 'Find Me button',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableLayoutToggle', {
                  label: 'Layout toggle (drill / vertical / horizontal)',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableStats', {
                  label: 'Org stats bar',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableDeptFilter', {
                  label: 'Department filter',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneToggle('enableUserFilter', {
                  label: 'User type filter (members / guests)',
                  onText: 'Visible',
                  offText: 'Hidden'
                })
              ]
            },
            {
              groupName: 'Directory',
              groupFields: [
                PropertyPaneSlider('pageSize', {
                  label: 'Max employees per page',
                  min: 10,
                  max: 200,
                  value: 50,
                  step: 10,
                  showValue: true
                }),
                PropertyPaneLabel('pageSize', {
                  text: 'Card size, visible fields, and other display preferences are set per-user via the Settings gear in the app.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
