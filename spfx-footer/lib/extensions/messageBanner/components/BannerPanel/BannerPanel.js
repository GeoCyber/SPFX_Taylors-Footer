import * as React from 'react';
import styles from './BannerPanel.module.scss';
import * as strings from 'MessageBannerApplicationCustomizerStrings';
import { PanelType, Panel } from 'office-ui-fabric-react/lib/Panel';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { SwatchColorPicker } from 'office-ui-fabric-react/lib/SwatchColorPicker';
import { getColorFromString } from 'office-ui-fabric-react/lib/Color';
import { Slider } from 'office-ui-fabric-react/lib/Slider';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
var TEXT_COLORS = [
    { id: 't1', label: 'Black', color: '#000000' },
    { id: 't2', label: 'Red', color: '#ff0000' },
    { id: 't3', label: 'Blue', color: '#1d45ba' },
    { id: 't4', label: 'White', color: '#ffffff' }
];
var BACKGROUND_COLORS = [
    { id: 'b1', label: 'Yellow', color: '#ffff00' },
    { id: 'b2', label: 'Light Yellow', color: '#ffffc6' },
    { id: 'b3', label: 'Teal', color: '#038387' },
    { id: 'b4', label: 'Blue', color: '#0078d4' },
    { id: 'b5', label: 'Dark Red', color: '#ba2a1d' },
    { id: 'b6', label: 'Salmon', color: '#e9967a' },
    { id: 'b7', label: 'Orange', color: '#ff8c00' },
    { id: 'b8', label: 'White', color: '#ffffff' }
];
var BannerPanel = function (props) {
    var textColor = getColorFromString(props.settings.textColor);
    var textColorMatch = TEXT_COLORS.filter(function (c) { return textColor && c.color === textColor.str; });
    var textColorSelectedId = textColorMatch && textColorMatch.length > 0 ? textColorMatch[0].id : null;
    var backgroundColor = getColorFromString(props.settings.backgroundColor);
    var backgroundColorMatch = BACKGROUND_COLORS.filter(function (c) { return backgroundColor && c.color === backgroundColor.str; });
    var backgroundColorSelectedId = backgroundColorMatch && backgroundColorMatch.length > 0 ? backgroundColorMatch[0].id : null;
    return (React.createElement(Panel, { isOpen: props.isOpen, isBlocking: false, isLightDismiss: true, type: PanelType.smallFixedFar, onDismiss: props.onCancelOrDismiss, headerText: strings.BannerPanelHeaderText, className: styles.BannerPanelContainer, onRenderFooterContent: function () { return (React.createElement("div", { className: styles.FooterButtons },
            React.createElement(PrimaryButton, { onClick: props.onSave, disabled: props.isSaving }, strings.BannerPanelButtonSaveText),
            React.createElement(DefaultButton, { onClick: props.onCancelOrDismiss, disabled: props.isSaving }, strings.BannerPanelButtonCancelText),
            props.isSaving && React.createElement(Spinner, { size: SpinnerSize.small }),
            React.createElement("div", { className: styles.ResetToDefaults, onClick: props.resetToDefaults }, strings.BannerPanelButtonResetToDefaultsText))); } },
        React.createElement("div", { className: styles.FieldContainer },
            React.createElement("div", { className: styles.FieldSection },
                React.createElement(Label, { className: styles.FieldLabel }, strings.BannerPanelFieldMessageLabel),
                React.createElement(Label, { className: styles.FieldDescription }, strings.BannerPanelFieldMessageDescription),
                React.createElement(TextField, { multiline: true, rows: 5, value: props.settings.message, className: styles.SwatchColorPicker, onChange: function (e, value) { return props.onFieldChange({ message: value }); } })),
            React.createElement("div", { className: styles.FieldSection },
                React.createElement(Label, { className: styles.FieldLabel }, strings.BannerPanelFieldTextColorLabel),
                React.createElement(SwatchColorPicker, { columnCount: 10, selectedId: textColorSelectedId, cellShape: 'circle', colorCells: TEXT_COLORS, className: styles.SwatchColorPicker, onColorChanged: function (e, value) { return props.onFieldChange({ textColor: value }); } }),
                React.createElement(TextField, { defaultValue: props.settings.textColor, onChange: function (e, value) { return props.onFieldChange({ textColor: value }); } })),
            React.createElement("div", { className: styles.FieldSection },
                React.createElement(Label, { className: styles.FieldLabel }, strings.BannerPanelFieldBackgroundColorLabel),
                React.createElement(SwatchColorPicker, { columnCount: 10, selectedId: backgroundColorSelectedId, cellShape: 'circle', colorCells: BACKGROUND_COLORS, className: styles.SwatchColorPicker, onColorChanged: function (e, value) { return props.onFieldChange({ backgroundColor: value }); } }),
                React.createElement(TextField, { defaultValue: props.settings.backgroundColor, onChange: function (e, value) { return props.onFieldChange({ backgroundColor: value }); } })),
            React.createElement("div", { className: styles.FieldSection },
                React.createElement(Label, { className: styles.FieldLabel }, strings.BannerPanelFieldTextSizeLabel),
                React.createElement(Slider, { min: 14, max: 50, step: 2, value: props.settings.textFontSizePx, showValue: true, onChange: function (value) { return props.onFieldChange({ textFontSizePx: value }); } })),
            React.createElement("div", { className: styles.FieldSection },
                React.createElement(Label, { className: styles.FieldLabel }, strings.BannerPanelFieldBannerHeightLabel),
                React.createElement(Slider, { min: 20, max: 80, step: 2, value: props.settings.bannerHeightPx, showValue: true, onChange: function (value) { return props.onFieldChange({ bannerHeightPx: value }); } })),
            React.createElement("div", { className: styles.FieldSection },
                React.createElement(Label, { className: styles.FieldLabel }, strings.BannerPanelFieldVisibleStartDateLabel),
                React.createElement(Toggle, { checked: props.settings.visibleStartDate !== null, onText: strings.BannerPanelFieldVisibleStartDateEnabledLabel, offText: strings.BannerPanelFieldVisibleStartDateDisabledLabel, onChange: function (ev, value) { return props.onFieldChange({ visibleStartDate: value ? new Date() : null }); } }),
                props.settings.visibleStartDate && (React.createElement(DatePicker, { value: new Date(props.settings.visibleStartDate), onSelectDate: function (value) { return props.onFieldChange({ visibleStartDate: value.toDateString() }); } }))))));
};
export default BannerPanel;
//# sourceMappingURL=BannerPanel.js.map