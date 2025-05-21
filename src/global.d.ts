/// <reference types="@types/google-apps-script" />

declare const PropertiesService: GoogleAppsScript.Properties.PropertiesService;
declare const SpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp;
declare const Logger: GoogleAppsScript.Base.Logger;

interface YouTube {
	Search: typeof GoogleAppsScript.YouTube.Search;
	Videos: typeof GoogleAppsScript.YouTube.Videos;
	Channels: typeof GoogleAppsScript.YouTube.Channels;
}

declare const YouTube: YouTube;
