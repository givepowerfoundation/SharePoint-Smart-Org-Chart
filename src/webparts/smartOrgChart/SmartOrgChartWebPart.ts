import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType,
  PropertyPaneTextField,
  PropertyPaneDropdown,
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
  defaultZoom: number;
  defaultFontScale: number;
  // Data
  dataSource: 'auto' | 'graph' | 'search';
  dottedLineAttribute: string;
  // User filters
  excludedAccounts: string;
  restrictToTenantDomain: boolean;
  hideGuestUsers: boolean;
  hideDisabledAccounts: boolean;
  hideNoJobTitle: boolean;
  hideNoDepartment: boolean;
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
    if (p.defaultZoom      === undefined) p.defaultZoom      = 0;
    if (p.defaultFontScale === undefined) p.defaultFontScale = 1;
    // Branding defaults
    if (p.companyName === undefined) p.companyName = '';
    if (p.logoUrl     === undefined) p.logoUrl     = '';
    // Data source default
    if (p.dataSource  === undefined) p.dataSource  = 'auto';
    if (p.dottedLineAttribute === undefined) p.dottedLineAttribute = '';
    // User filter defaults
    if (p.excludedAccounts       === undefined) p.excludedAccounts       = '';
    if (p.restrictToTenantDomain === undefined) p.restrictToTenantDomain = false;
    if (p.hideGuestUsers         === undefined) p.hideGuestUsers         = true;
    if (p.hideDisabledAccounts   === undefined) p.hideDisabledAccounts   = true;
    if (p.hideNoJobTitle         === undefined) p.hideNoJobTitle         = false;
    if (p.hideNoDepartment       === undefined) p.hideNoDepartment       = false;
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
      defaultZoom: this.properties.defaultZoom ?? 0,
      defaultFontScale: this.properties.defaultFontScale || 1,
      enableFindMe: this.properties.enableFindMe !== false,
      enableLayoutToggle: this.properties.enableLayoutToggle !== false,
      enableStats: this.properties.enableStats !== false,
      enableDeptFilter: this.properties.enableDeptFilter !== false,
      enableUserFilter: this.properties.enableUserFilter !== false,
      dataSource: this.properties.dataSource || 'auto',
      dottedLineAttribute: this.properties.dottedLineAttribute || '',
      excludedAccounts:       this.properties.excludedAccounts       || '',
      restrictToTenantDomain: this.properties.restrictToTenantDomain || false,
      hideGuestUsers:         this.properties.hideGuestUsers         ?? true,
      hideDisabledAccounts:   this.properties.hideDisabledAccounts   ?? true,
      hideNoJobTitle:         this.properties.hideNoJobTitle         || false,
      hideNoDepartment:       this.properties.hideNoDepartment       || false,
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
                PropertyPaneDropdown('defaultFontScale', {
                  label: 'Default Font Size',
                  options: [
                    { key: 0.75, text: '75% — Extra Small' },
                    { key: 0.85, text: '85% — Small' },
                    { key: 1,    text: '100% — Normal' },
                    { key: 1.15, text: '115% — Large' },
                    { key: 1.3,  text: '130% — Extra Large' },
                    { key: 1.5,  text: '150% — XXL' },
                    { key: 1.75, text: '175% — XXXL' },
                  ],
                  selectedKey: this.properties.defaultFontScale || 1,
                }),
                PropertyPaneChoiceGroup('defaultLayout', {
                  label: 'Default Org Chart Layout',
                  options: [
                    { key: 'drill',      text: 'Drill-Down',    iconProps: { officeFabricIconFontName: 'Org' } },
                    { key: 'vertical',   text: 'Top Down',      iconProps: { officeFabricIconFontName: 'Down' } },
                    { key: 'horizontal', text: 'Left to Right', iconProps: { officeFabricIconFontName: 'Forward' } },
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
                PropertyPaneLabel('hideDisabledAccounts', {
                  text: 'Note: guest and disabled account detection requires the Graph API data source. SharePoint Search does not return this information, so these two filters (and the Guest/Disabled badges) have no effect when data comes from Search.'
                }),
                PropertyPaneToggle('hideNoJobTitle', {
                  label: 'Hide accounts without a job title',
                  onText: 'On — accounts with no job title hidden',
                  offText: 'Off — all accounts shown regardless of job title'
                }),
                PropertyPaneToggle('hideNoDepartment', {
                  label: 'Hide accounts without a department',
                  onText: 'On — accounts with no department hidden',
                  offText: 'Off — all accounts shown regardless of department'
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
                }),
                PropertyPaneTextField('dottedLineAttribute', {
                  label: 'Dotted-line manager attribute',
                  placeholder: 'extensionAttribute10',
                  description: 'Optional. Name of the Azure AD on-premises extension attribute (extensionAttribute1-15) that stores a secondary "dotted line" manager\'s email or UPN. Dotted-line relationships appear on profile cards. Requires the Graph API data source.'
                }),
                PropertyPaneDropdown('defaultZoom', {
                  label: 'Default Org Chart Zoom',
                  options: [
                    { key: 0,    text: 'Default (Auto-fit)' },
                    { key: 0.5,  text: '50%' },
                    { key: 0.75, text: '75%' },
                    { key: 1,    text: '100%' },
                    { key: 1.25, text: '125%' },
                    { key: 1.5,  text: '150%' },
                  ],
                  selectedKey: this.properties.defaultZoom ?? 0,
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
                ({
                  type: PropertyPaneFieldType.Custom,
                  targetProperty: 'pageSize',
                  properties: {
                    key: 'pageSizeField',
                    onRender: (elem: HTMLElement, _ctx: any, changeCallback: ((targetProperty?: string, newValue?: any) => void) | undefined) => {
                    elem.innerHTML = '';
                    const current = this.properties.pageSize || 50;

                    const label = document.createElement('label');
                    label.textContent = 'Max employees per page';
                    label.style.cssText = 'display:block;font-weight:600;font-size:14px;color:#323130;margin-bottom:8px;';

                    const row = document.createElement('div');
                    row.style.cssText = 'display:flex;align-items:center;gap:12px;';

                    const slider = document.createElement('input');
                    slider.type = 'range';
                    slider.min = '10';
                    slider.max = '200';
                    slider.step = '1';
                    slider.value = String(current);
                    slider.style.cssText = 'flex:1;accent-color:#0078d4;';

                    const numInput = document.createElement('input');
                    numInput.type = 'number';
                    numInput.min = '10';
                    numInput.max = '200';
                    numInput.value = String(current);
                    numInput.style.cssText = 'width:64px;padding:4px 6px;border:1px solid #c8c6c4;border-radius:2px;font-size:14px;text-align:center;';

                    const sync = (val: number): void => {
                      const v = Math.max(10, Math.min(200, isNaN(val) ? 50 : val));
                      slider.value = String(v);
                      numInput.value = String(v);
                      if (changeCallback) changeCallback('pageSize', v);
                    };

                    slider.addEventListener('input', () => sync(parseInt(slider.value, 10)));
                    numInput.addEventListener('change', () => sync(parseInt(numInput.value, 10)));

                    row.appendChild(slider);
                    row.appendChild(numInput);
                    elem.appendChild(label);
                    elem.appendChild(row);
                  },
                    onDispose: (elem: HTMLElement) => { elem.innerHTML = ''; }
                  }
                } as IPropertyPaneField<IPropertyPaneCustomFieldProps>),
                PropertyPaneLabel('pageSize', {
                  text: 'Card size, visible fields, and other display preferences are set per-user via the Settings gear in the app.'
                })
              ]
            },
            {
              groupName: 'Demo',
              groupFields: [
                PropertyPaneToggle('useDemoData', {
                  label: 'Use Demo Data',
                  onText: 'On — showing sample employees',
                  offText: 'Off — live Microsoft 365 data',
                  checked: false
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
