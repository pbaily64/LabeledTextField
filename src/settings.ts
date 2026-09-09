"use strict";

import powerbi from "powerbi-visuals-api";
import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard  = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

// ── Carte "Style du libellé" ────────────────────────────────────────────────
export class LabelStyleCard extends FormattingSettingsCard {

    // Texte / fx — activé ConstantOrRule pour pouvoir y brancher une mesure DAX
    titleText = new formattingSettings.TextInput({
        name: "titleText",
        displayName: "Title of the text",
        value: "",
        placeholder: "Title…",
        instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        value: { value: "#FFFFFF" },
        instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background color",
        value: { value: "#E66C37" },
        instanceKind: powerbi.VisualEnumerationInstanceKinds.ConstantOrRule
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font size",
        value: 12,
        options: {
            minValue: { value: 8,  type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 40, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    fontBold = new formattingSettings.ToggleSwitch({
        name: "fontBold",
        displayName: "Bold",
        value: true
    });

    minWidth = new formattingSettings.NumUpDown({
        name: "minWidth",
        displayName: "Minimum width (px)",
        value: 120,
        options: {
            minValue: { value: 40,  type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 400, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    paddingH = new formattingSettings.NumUpDown({
        name: "paddingH",
        displayName: "Horizontal padding (px)",
        value: 10,
        options: {
            minValue: { value: 0,  type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 40, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    borderRadius = new formattingSettings.NumUpDown({
        name: "borderRadius",
        displayName: "Border radius (px)",
        value: 4,
        options: {
            minValue: { value: 0,  type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 20, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    name: string        = "labelStyle";
    displayName: string = "Label style";
    slices: Array<FormattingSettingsSlice> = [
        this.titleText,
        this.fontColor,
        this.backgroundColor,
        this.fontSize,
        this.fontBold,
        this.minWidth,
        this.paddingH,
        this.borderRadius
    ];
}

// ── Carte "Data style" (ex "Area style") ────────────────────────────────────
// NOTE : le name interne reste "fieldStyle" pour ne pas invalider les
// propriétés déjà enregistrées dans les rapports existants.
export class FieldStyleCard extends FormattingSettingsCard {

    displayAsCheckbox = new formattingSettings.ToggleSwitch({
        name: "displayAsCheckbox",
        displayName: "Display as checkbox",
        value: false
    });

    checkboxColor = new formattingSettings.ColorPicker({
        name: "checkboxColor",
        displayName: "checkbox color (checked)",
        value: { value: "#0078D4" }
    });

    checkboxSize = new formattingSettings.NumUpDown({
        name: "checkboxSize",
        displayName: "Checkbox size (px)",
        value: 16,
        options: {
            minValue: { value: 10, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 40, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "Background color",
        value: { value: "#FAFAFA" }
    });

    borderColor = new formattingSettings.ColorPicker({
        name: "borderColor",
        displayName: "Border color",
        value: { value: "#CCCCCC" }
    });

    borderWidth = new formattingSettings.NumUpDown({
        name: "borderWidth",
        displayName: "Border width (px)",
        value: 1,
        options: {
            minValue: { value: 0, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 6, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        value: { value: "#252423" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Font size",
        value: 12,
        options: {
            minValue: { value: 8,  type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 40, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    fontBold = new formattingSettings.ToggleSwitch({
        name: "fontBold",
        displayName: "Bold",
        value: false
    });

    paddingH = new formattingSettings.NumUpDown({
        name: "paddingH",
        displayName: "Horizontal padding (px)",
        value: 8,
        options: {
            minValue: { value: 0,  type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 40, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    // ── Format d'affichage de la valeur ──────────────────────────────────────
    // "none" = comportement historique (auto-détection via source.format)
    textFormat = new formattingSettings.AutoDropdown({
        name: "textFormat",
        displayName: "Text format",
        value: "none"
    });

    // Nombre de décimales — visible uniquement pour les formats numériques
    // (visibilité pilotée dans visual.getFormattingModel)
    decimals = new formattingSettings.NumUpDown({
        name: "decimals",
        displayName: "Decimals",
        value: 2,
        options: {
            minValue: { value: 0, type: powerbi.visuals.ValidatorType.Min },
            maxValue: { value: 6, type: powerbi.visuals.ValidatorType.Max }
        }
    });

    // Position du symbole monétaire — visible uniquement pour les formats
    // Currency (visibilité pilotée dans visual.getFormattingModel)
    symbolPosition = new formattingSettings.AutoDropdown({
        name: "symbolPosition",
        displayName: "Symbol position",
        value: "auto"
    });

    // Masque de date/heure (syntaxe .NET) — visible uniquement pour le
    // format Date (visibilité pilotée dans visual.getFormattingModel)
    datePattern = new formattingSettings.TextInput({
        name: "datePattern",
        displayName: "Date pattern",
        value: "dd/MM/yyyy",
        placeholder: "dd/MM/yyyy"
    });

    alignment = new formattingSettings.AutoDropdown({
        name: "alignment",
        displayName: "Text alignment",
        value: "left"
    });

    placeholderText = new formattingSettings.TextInput({
        name: "placeholderText",
        displayName: "Text if empty",
        value: "",
        placeholder: "(empty)"
    });

    name: string        = "fieldStyle";
    displayName: string = "Data style";
    slices: Array<FormattingSettingsSlice> = [
        this.displayAsCheckbox,
        this.checkboxColor,
        this.checkboxSize,
        this.backgroundColor,
        this.borderColor,
        this.borderWidth,
        this.fontColor,
        this.fontSize,
        this.fontBold,
        this.paddingH,
        this.textFormat,
        this.decimals,
        this.symbolPosition,
        this.datePattern,
        this.alignment,
        this.placeholderText
    ];
}

// ── Modèle global ───────────────────────────────────────────────────────────
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    labelStyle = new LabelStyleCard();
    fieldStyle = new FieldStyleCard();
    cards      = [this.labelStyle, this.fieldStyle];
}
