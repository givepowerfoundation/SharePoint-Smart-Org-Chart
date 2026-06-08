import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Slider } from '@fluentui/react/lib/Slider';
import { Separator } from '@fluentui/react/lib/Separator';
import { Label } from '@fluentui/react/lib/Label';
import { IUserSettings } from '../ISmartOrgChartProps';
import { ISettingsPanelProps } from './ISettingsPanelProps';
import styles from './SettingsPanel.module.scss';

interface ISettingsPanelState {
  draft: IUserSettings;
}

const alphabetOptions: IDropdownOption[] = [
  { key: 'firstName', text: 'First Name' },
  { key: 'lastName', text: 'Last Name' }
];

const cardSizeOptions: IDropdownOption[] = [
  { key: 'small', text: 'Small — compact, more cards visible' },
  { key: 'medium', text: 'Medium — balanced' },
  { key: 'large', text: 'Large — spacious, fewer cards' }
];

const fontScaleOptions: IDropdownOption[] = [
  { key: 0.75, text: '75% — Extra Small' },
  { key: 0.85, text: '85% — Small' },
  { key: 1,    text: '100% — Normal' },
  { key: 1.15, text: '115% — Large' },
  { key: 1.3,  text: '130% — Extra Large' },
  { key: 1.5,  text: '150% — XXL' },
  { key: 1.75, text: '175% — XXXL' },
];

export class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState> {
  constructor(props: ISettingsPanelProps) {
    super(props);
    this.state = { draft: { ...props.settings } };
  }

  public componentDidUpdate(prev: ISettingsPanelProps): void {
    if (!prev.isOpen && this.props.isOpen) {
      this.setState({ draft: { ...this.props.settings } });
    }
  }

  private _update = <K extends keyof IUserSettings>(key: K, value: IUserSettings[K]): void => {
    this.setState(prev => ({ draft: { ...prev.draft, [key]: value } }));
  }

  private _save = (): void => {
    this.props.onSave(this.state.draft);
  }

  private _reset = (): void => {
    this.setState({ draft: { ...this.props.settings } });
  }

  public render(): React.ReactElement {
    const { isOpen, onDismiss, mockSize, onMockSizeChange } = this.props;
    const { draft } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={onDismiss}
        type={PanelType.medium}
        headerText="Preferences"
        isFooterAtBottom
        onRenderFooterContent={() => (
          <div className={styles.footer}>
            <PrimaryButton text="Save" onClick={this._save} />
            <DefaultButton text="Cancel" onClick={onDismiss} style={{ marginLeft: 8 }} />
            <DefaultButton text="Discard Changes" onClick={this._reset} style={{ marginLeft: 'auto' }} />
          </div>
        )}
      >
        <div className={styles.body}>

          {/* ── Directory display ── */}
          <Separator alignContent="start">
            <span className={styles.sectionLabel}>Directory</span>
          </Separator>

          <Dropdown
            label="Alphabet Filter By"
            selectedKey={draft.alphabetFilterField}
            options={alphabetOptions}
            onChange={(_, o) => o && this._update('alphabetFilterField', o.key as 'firstName' | 'lastName')}
          />

          <Dropdown
            label="Card Size"
            selectedKey={draft.cardSize}
            options={cardSizeOptions}
            onChange={(_, o) => o && this._update('cardSize', o.key as 'small' | 'medium' | 'large')}
            className={styles.field}
          />

          <Dropdown
            label="Font Size"
            selectedKey={draft.fontScale || 1}
            options={fontScaleOptions}
            onChange={(_, o) => o && this._update('fontScale', o.key as number)}
            className={styles.field}
          />

          {/* ── Cards & Fields ── */}
          <Separator alignContent="start" className={styles.separator}>
            <span className={styles.sectionLabel}>Cards &amp; Fields</span>
          </Separator>

          <div className={styles.toggleGroup}>
            <Label>Show on employee cards:</Label>
            <Toggle
              label="Email address"
              checked={draft.showEmail}
              onChange={(_, v) => this._update('showEmail', !!v)}
              inlineLabel
            />
            <Toggle
              label="Phone number"
              checked={draft.showPhone}
              onChange={(_, v) => this._update('showPhone', !!v)}
              inlineLabel
            />
            <Toggle
              label="Department"
              checked={draft.showDepartment}
              onChange={(_, v) => this._update('showDepartment', !!v)}
              inlineLabel
            />
            <Toggle
              label="Office location"
              checked={draft.showOffice}
              onChange={(_, v) => this._update('showOffice', !!v)}
              inlineLabel
            />
          </div>

          {/* ── Org Chart ── */}
          <Separator alignContent="start" className={styles.separator}>
            <span className={styles.sectionLabel}>Org Chart</span>
          </Separator>

          <Slider
            label={`Manager levels shown when focusing a person: ${draft.levelsAbove}`}
            min={0}
            max={5}
            step={1}
            value={draft.levelsAbove}
            onChange={v => this._update('levelsAbove', v)}
            className={styles.field}
            showValue={false}
          />

          <Toggle
            label="Compact cards"
            checked={draft.compactCards}
            onChange={(_, v) => this._update('compactCards', !!v)}
            inlineLabel
            className={styles.field}
            onText="On — smaller cards, more visible at once"
            offText="Off — standard size"
          />

          <div className={styles.hint}>
            These preferences are saved to your browser and apply only to you.
            Chart theme, default view, and layout are configured by your SharePoint admin.
          </div>

          {/* ── Demo data ── */}
          {mockSize !== undefined && onMockSizeChange && (
            <>
              <Separator alignContent="start" className={styles.separator}>
                <span className={styles.sectionLabel}>Demo Data</span>
              </Separator>
              <Label>Dataset size</Label>
              <div className={styles.mockSizeBtns}>
                {([150, 500, 1000] as const).map(s => (
                  <button
                    key={s}
                    className={`${styles.mockSizeBtn}${mockSize === s ? ` ${styles.mockSizeBtnActive}` : ''}`}
                    onClick={() => onMockSizeChange(s)}
                  >
                    {s.toLocaleString()} people
                  </button>
                ))}
              </div>
            </>
          )}
        </div>
      </Panel>
    );
  }
}
