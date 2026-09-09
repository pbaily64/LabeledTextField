"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions       = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual                   = powerbi.extensibility.visual.IVisual;
import IVisualHost               = powerbi.extensibility.visual.IVisualHost;
import DataView                  = powerbi.DataView;

import { VisualFormattingSettingsModel } from "./settings";

// État booléen résolu à partir de la donnée brute
type BoolState = "true" | "false" | "unknown";

// Formats numériques pour lesquels le slice "Decimals" est pertinent
const NUMERIC_FORMATS = ["decimalNumber", "percentage", "currencyEUR", "currencyUSD"];

// Formats monétaires pour lesquels le slice "Symbol position" est pertinent
const CURRENCY_FORMATS = ["currencyEUR", "currencyUSD"];

export class Visual implements IVisual {

    private host:      IVisualHost;
    private container: HTMLDivElement;
    private labelEl:   HTMLSpanElement;
    private valueEl:   HTMLSpanElement;

    // Locale transmise par Power BI (ex. "fr-BE", "nl-BE", "en-US").
    // Fallback fr-BE si l'hôte ne la fournit pas.
    private locale: string = "fr-BE";

    private formattingSettings:        VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private dataView: DataView | null = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();
        this.locale = options.host.locale || "fr-BE";

        this.container = document.createElement("div");
        this.container.className = "ltf-container";

        this.labelEl = document.createElement("span");
        this.labelEl.className = "ltf-label";

        this.valueEl = document.createElement("span");
        this.valueEl.className = "ltf-value";

        this.container.appendChild(this.labelEl);
        this.container.appendChild(this.valueEl);
        options.element.appendChild(this.container);
    }

    public update(options: VisualUpdateOptions): void {
        if (!options?.dataViews?.[0]) return;

        this.dataView = options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(
            VisualFormattingSettingsModel,
            this.dataView
        );

        this.applyFormatting();
        this.renderLabel();
        this.renderValue();
    }

    // ── Résolution du libellé ────────────────────────────────────────────────
    // Priorité 1 : bucket "Libellé (DAX)"  → mesure dans categorical.values
    // Priorité 2 : propriété titleText du Format pane (constante ou fx)
    // Priorité 3 : nom d'affichage du champ branché dans "Valeur"
    private resolveLabelText(): string {
        const cat = this.dataView?.categorical;

        if (cat?.values) {
            for (const col of cat.values) {
                if (col.source?.roles?.["labelMeasure"]) {
                    const v = col.values?.[0];
                    if (v != null && String(v).trim() !== "") return String(v);
                }
            }
        }

        const paneText = this.formattingSettings.labelStyle.titleText.value?.trim();
        if (paneText) return paneText;

        if (cat?.categories?.[0]?.source?.displayName) {
            return cat.categories[0].source.displayName;
        }
        if (cat?.values) {
            for (const col of cat.values) {
                if (!col.source?.roles?.["labelMeasure"]) {
                    return col.source.displayName ?? "";
                }
            }
        }
        return "";
    }

    // ── Récupère la valeur brute du champ (sans formatage) ───────────────────
    private getRawFieldValue(): { raw: unknown; source: powerbi.DataViewMetadataColumn | undefined } {
        const cat = this.dataView?.categorical;
        if (!cat) return { raw: null, source: undefined };

        if (cat.categories?.[0]) {
            const vals = cat.categories[0].values;
            if (vals?.length) {
                return { raw: vals[0], source: cat.categories[0].source };
            }
        }

        if (cat.values) {
            for (const col of cat.values) {
                if (col.source?.roles?.["labelMeasure"]) continue;
                const raw = col.values?.[0];
                if (raw !== null && raw !== undefined) {
                    return { raw, source: col.source };
                }
            }
        }
        return { raw: null, source: undefined };
    }

    // ── Interprétation booléenne d'une valeur ────────────────────────────────
    // Power BI peut remonter : true/false (bool natif), "True"/"False" (string
    // DAX), 1/0, "1"/"0", "oui"/"non", "yes"/"no", "vrai"/"faux"…
    private interpretBoolean(raw: unknown): BoolState {
        if (raw === null || raw === undefined) return "unknown";
        if (typeof raw === "boolean") return raw ? "true" : "false";
        if (typeof raw === "number") {
            if (raw === 1) return "true";
            if (raw === 0) return "false";
            return "unknown";
        }
        const s = String(raw).trim().toLowerCase();
        if (!s) return "unknown";
        if (["true", "1", "oui", "yes", "vrai", "y", "o"].indexOf(s) >= 0) return "true";
        if (["false", "0", "non", "no", "faux", "n"].indexOf(s) >= 0) return "false";
        return "unknown";
    }

    // ── Conversion sûre vers number / Date ───────────────────────────────────
    private toNumber(raw: unknown): number | null {
        if (typeof raw === "number") return isFinite(raw) ? raw : null;
        if (typeof raw === "string") {
            // Tolère la virgule décimale ("1234,56") remontée en texte
            const n = parseFloat(raw.replace(",", "."));
            return isNaN(n) ? null : n;
        }
        if (raw instanceof Date) return raw.getTime();
        return null;
    }

    private toDate(raw: unknown): Date | null {
        if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;
        if (typeof raw === "string" || typeof raw === "number") {
            const d = new Date(raw);
            return isNaN(d.getTime()) ? null : d;
        }
        return null;
    }

    // ── Moteur de format explicite (propriété "Text format") ─────────────────
    // Retourne null si la valeur n'est pas convertible dans le format demandé,
    // auquel cas l'appelant retombe sur le formatage automatique.
    private applyExplicitFormat(raw: unknown, fmt: string, decimals: number): string | null {
        const L = this.locale;

        switch (fmt) {
            case "wholeNumber": {
                const n = this.toNumber(raw);
                if (n === null) return null;
                return n.toLocaleString(L, { maximumFractionDigits: 0 });
            }
            case "decimalNumber": {
                const n = this.toNumber(raw);
                if (n === null) return null;
                return n.toLocaleString(L, {
                    minimumFractionDigits: decimals,
                    maximumFractionDigits: decimals
                });
            }
            case "percentage": {
                const n = this.toNumber(raw);
                if (n === null) return null;
                // style "percent" multiplie par 100 et place le % selon la locale
                return n.toLocaleString(L, {
                    style: "percent",
                    minimumFractionDigits: decimals,
                    maximumFractionDigits: decimals
                });
            }
            case "currencyEUR":
            case "currencyUSD": {
                const n = this.toNumber(raw);
                if (n === null) return null;

                const position = (this.formattingSettings.fieldStyle.symbolPosition.value as string) || "auto";

                // Auto : Intl place le symbole selon les conventions de la locale
                if (position === "auto") {
                    return n.toLocaleString(L, {
                        style: "currency",
                        currency: fmt === "currencyEUR" ? "EUR" : "USD",
                        minimumFractionDigits: decimals,
                        maximumFractionDigits: decimals
                    });
                }

                // Before / After : nombre formaté par la locale (séparateurs
                // corrects), symbole placé manuellement avec espace insécable
                const symbol = fmt === "currencyEUR" ? "\u20AC" : "$";
                const num = n.toLocaleString(L, {
                    minimumFractionDigits: decimals,
                    maximumFractionDigits: decimals
                });
                return position === "before"
                    ? symbol + "\u00A0" + num
                    : num + "\u00A0" + symbol;
            }
            case "date": {
                const d = this.toDate(raw);
                if (d === null) return null;
                const pattern = (this.formattingSettings.fieldStyle.datePattern.value as string) || "dd/MM/yyyy";
                return this.formatDatePattern(d, pattern);
            }
            default:
                return null;
        }
    }

    // ── Formatage de date par masque (syntaxe .NET, cf. FORMAT() en DAX) ─────
    // Jetons supportés : dddd ddd dd d | MMMM MMM MM M | yyyy yy
    //                    HH H hh h | mm | ss | tt (AM/PM)
    // Les noms de mois/jours suivent host.locale (fr → juillet, nl → juli…).
    // Tout autre caractère est recopié tel quel.
    // IMPORTANT : remplacement en une seule passe via replace(regex, callback)
    // pour éviter que les jetons ne re-matchent le texte déjà produit
    // (ex. le "d" de "jeudi").
    private formatDatePattern(d: Date, pattern: string): string {
        const L = this.locale;
        const pad = (n: number): string => String(n).padStart(2, "0");

        const h24 = d.getHours();
        const h12 = h24 % 12 === 0 ? 12 : h24 % 12;

        const tokenRegex = /dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy|HH|H|hh|h|mm|ss|tt/g;

        return pattern.replace(tokenRegex, (token: string): string => {
            switch (token) {
                case "dddd": return d.toLocaleDateString(L, { weekday: "long"  });
                case "ddd":  return d.toLocaleDateString(L, { weekday: "short" });
                case "dd":   return pad(d.getDate());
                case "d":    return String(d.getDate());
                case "MMMM": return d.toLocaleDateString(L, { month: "long"  });
                case "MMM":  return d.toLocaleDateString(L, { month: "short" });
                case "MM":   return pad(d.getMonth() + 1);
                case "M":    return String(d.getMonth() + 1);
                case "yyyy": return String(d.getFullYear());
                case "yy":   return pad(d.getFullYear() % 100);
                case "HH":   return pad(h24);
                case "H":    return String(h24);
                case "hh":   return pad(h12);
                case "h":    return String(h12);
                case "mm":   return pad(d.getMinutes());
                case "ss":   return pad(d.getSeconds());
                case "tt":   return h24 < 12 ? "AM" : "PM";
                default:     return token;
            }
        });
    }

    // ── Formatage automatique (comportement historique, format "None") ───────
    private formatValueAuto(raw: unknown, source: powerbi.DataViewMetadataColumn | undefined): string {
        if (raw === null || raw === undefined) return "";

        if (raw instanceof Date) {
            return raw.toLocaleDateString(this.locale);
        }

        if (typeof raw === "number") {
            const fmt = source?.format || "";
            try {
                if (fmt.includes("%")) {
                    return (raw * 100).toLocaleString(this.locale, { maximumFractionDigits: 2 }) + "%";
                }
                if (fmt.includes("€")) {
                    return raw.toLocaleString(this.locale, { style: "currency", currency: "EUR" });
                }
                if (fmt.includes("$")) {
                    return raw.toLocaleString(this.locale, { style: "currency", currency: "USD" });
                }
            } catch { /* ignore */ }
            return raw.toLocaleString(this.locale);
        }

        // Booléens natifs en mode texte : on garde True/False explicite
        if (typeof raw === "boolean") return raw ? "True" : "False";

        return String(raw);
    }

    // ── Point d'entrée du formatage texte ────────────────────────────────────
    private formatValue(raw: unknown, source: powerbi.DataViewMetadataColumn | undefined): string {
        if (raw === null || raw === undefined) return "";

        const fld      = this.formattingSettings.fieldStyle;
        const fmt      = (fld.textFormat.value as string) || "none";
        const decimals = Math.max(0, Math.min(6, (fld.decimals.value as number) ?? 2));

        if (fmt !== "none") {
            const formatted = this.applyExplicitFormat(raw, fmt, decimals);
            if (formatted !== null) return formatted;
            // Valeur non convertible dans le format demandé → fallback auto
        }

        return this.formatValueAuto(raw, source);
    }

    // ── Application du formatage visuel ──────────────────────────────────────
    private applyFormatting(): void {
        const lbl = this.formattingSettings.labelStyle;
        const fld = this.formattingSettings.fieldStyle;

        // Libellé
        const lblFc     = lbl.fontColor?.value?.value        || "#FFFFFF";
        const lblBg     = lbl.backgroundColor?.value?.value  || "#E66C37";
        const lblFs     = (lbl.fontSize.value as number)     ?? 12;
        const lblBold   = lbl.fontBold.value as boolean;
        const lblMinW   = (lbl.minWidth.value as number)     ?? 120;
        const lblPadH   = (lbl.paddingH.value as number)     ?? 10;
        const lblRadius = (lbl.borderRadius.value as number) ?? 4;

        this.labelEl.style.color           = lblFc;
        this.labelEl.style.backgroundColor = lblBg;
        this.labelEl.style.fontSize        = `${lblFs}px`;
        this.labelEl.style.fontWeight      = lblBold ? "bold" : "normal";
        this.labelEl.style.minWidth        = `${lblMinW}px`;
        this.labelEl.style.padding         = `0 ${lblPadH}px`;
        this.labelEl.style.borderRadius    = `${lblRadius}px 0 0 ${lblRadius}px`;

        // Zone de valeur
        const fldBg     = fld.backgroundColor?.value?.value || "#FAFAFA";
        const fldBd     = fld.borderColor?.value?.value     || "#CCCCCC";
        const fldBdW    = (fld.borderWidth.value as number) ?? 1;
        const fldFc     = fld.fontColor?.value?.value       || "#252423";
        const fldFs     = (fld.fontSize.value as number)    ?? 12;
        const fldBold   = fld.fontBold.value as boolean;
        const fldPadH   = (fld.paddingH.value as number)    ?? 8;
        const fldAlign  = (fld.alignment.value as string)   || "left";
        const asCheckbox= fld.displayAsCheckbox.value as boolean;

        this.valueEl.style.backgroundColor = fldBg;
        this.valueEl.style.border          = `${fldBdW}px solid ${fldBd}`;
        this.valueEl.style.borderLeft      = "none";
        this.valueEl.style.color           = fldFc;
        this.valueEl.style.fontSize        = `${fldFs}px`;
        this.valueEl.style.fontWeight      = fldBold ? "bold" : "normal";
        this.valueEl.style.padding         = `0 ${fldPadH}px`;
        this.valueEl.style.borderRadius    = `0 ${lblRadius}px ${lblRadius}px 0`;

        // En mode checkbox, la case est centrée par défaut sauf si l'user a
        // explicitement choisi un alignement
        if (asCheckbox && fldAlign === "left") {
            this.valueEl.style.textAlign = "center";
        } else {
            this.valueEl.style.textAlign = fldAlign;
        }
    }

    private renderLabel(): void {
        this.labelEl.textContent = this.resolveLabelText();
    }

    // ── Vide la zone de valeur ───────────────────────────────────────────────
    private clearValue(): void {
        while (this.valueEl.firstChild) {
            this.valueEl.removeChild(this.valueEl.firstChild);
        }
    }

    // ── Rendu de la valeur (texte OU checkbox selon le mode) ─────────────────
    private renderValue(): void {
        const fld          = this.formattingSettings.fieldStyle;
        const asCheckbox   = fld.displayAsCheckbox.value as boolean;
        const { raw, source } = this.getRawFieldValue();

        this.clearValue();

        if (asCheckbox) {
            const state = this.interpretBoolean(raw);
            const size  = (fld.checkboxSize.value as number) ?? 16;
            const color = fld.checkboxColor?.value?.value || "#0078D4";

            if (state === "unknown") {
                // Valeur non interprétable → fallback texte (placeholder ou
                // valeur brute formatée)
                const txt = (raw === null || raw === undefined)
                    ? (fld.placeholderText.value || "")
                    : this.formatValue(raw, source);
                this.valueEl.textContent = txt;
                this.valueEl.style.opacity = txt && raw == null ? "0.6" : "1";
                this.valueEl.title = txt;
                return;
            }

            const svg = this.buildCheckboxSvg(state === "true", size, color);
            this.valueEl.appendChild(svg);
            this.valueEl.style.opacity = "1";
            this.valueEl.title = state === "true" ? "True" : "False";
            return;
        }

        // Mode texte (comportement historique)
        if (raw === null || raw === undefined) {
            const ph = fld.placeholderText.value || "";
            this.valueEl.textContent = ph;
            this.valueEl.style.opacity = ph ? "0.6" : "1";
            this.valueEl.title = ph;
            return;
        }

        const v = this.formatValue(raw, source);
        this.valueEl.textContent = v;
        this.valueEl.style.opacity = "1";
        this.valueEl.title = v;
    }

    // ── Construit une checkbox SVG (cochée ou non) ───────────────────────────
    // Rendu manuel via createElementNS (conforme aux règles ESLint du projet,
    // cf. LabeledDateFilter.buildEraser)
    private buildCheckboxSvg(checked: boolean, size: number, color: string): SVGSVGElement {
        const svgNS = "http://www.w3.org/2000/svg";
        const svg   = document.createElementNS(svgNS, "svg");
        svg.setAttribute("width",   String(size));
        svg.setAttribute("height",  String(size));
        svg.setAttribute("viewBox", "0 0 24 24");
        svg.setAttribute("fill",    "none");

        const r = 3;  // rayon du coin
        const strokeW = 2;

        // Carré extérieur
        const rect = document.createElementNS(svgNS, "rect");
        rect.setAttribute("x",      String(strokeW / 2));
        rect.setAttribute("y",      String(strokeW / 2));
        rect.setAttribute("width",  String(24 - strokeW));
        rect.setAttribute("height", String(24 - strokeW));
        rect.setAttribute("rx",     String(r));
        rect.setAttribute("ry",     String(r));
        rect.setAttribute("stroke", checked ? color : "#888888");
        rect.setAttribute("stroke-width", String(strokeW));
        rect.setAttribute("fill",   checked ? color : "#FFFFFF");
        svg.appendChild(rect);

        if (checked) {
            // Marque de validation (✓)
            const check = document.createElementNS(svgNS, "path");
            check.setAttribute("d", "M6 12.5 L10.5 17 L18 8");
            check.setAttribute("stroke", "#FFFFFF");
            check.setAttribute("stroke-width", "2.5");
            check.setAttribute("stroke-linecap",  "round");
            check.setAttribute("stroke-linejoin", "round");
            check.setAttribute("fill", "none");
            svg.appendChild(check);
        }

        return svg;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        // Visibilité conditionnelle : "Decimals" n'apparaît que pour les
        // formats numériques (Decimal number, Percentage, Currency)
        const fld = this.formattingSettings?.fieldStyle;
        if (fld) {
            const fmt = (fld.textFormat.value as string) || "none";
            fld.decimals.visible       = NUMERIC_FORMATS.indexOf(fmt)  >= 0;
            fld.symbolPosition.visible = CURRENCY_FORMATS.indexOf(fmt) >= 0;
            fld.datePattern.visible    = fmt === "date";
        }
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
