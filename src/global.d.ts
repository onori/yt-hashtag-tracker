/// <reference types="@types/google-apps-script" />

declare const PropertiesService: GoogleAppsScript.Properties.PropertiesService;
declare const SpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp;
declare const Logger: GoogleAppsScript.Base.Logger;

// Extend the YouTubeAnalytics namespace
declare namespace GoogleAppsScript.YouTubeAnalytics {
  interface YouTubeAnalytics {
    Reports: {
      query: (params: {
        ids: string;
        startDate: string;
        endDate: string;
        metrics: string;
        dimensions?: string;
        filters?: string;
        sort?: string;
        maxResults?: number;
      }) => GoogleAppsScript.YouTubeAnalytics.Schema.QueryResponse;
    };
  }

  // Extend the Schema to include the QueryResponse interface
  namespace Schema {
    interface QueryResponse {
      columnHeaders?: Array<{
        name: string;
        columnType: string;
        dataType: string;
      }>;
      rows?: Array<Array<string | number | boolean | null>>;
    }
  }
}

// Declare the global YouTubeAnalytics object
declare const YouTubeAnalytics: GoogleAppsScript.YouTubeAnalytics.YouTubeAnalytics;

// Extend the global YouTube object
interface YouTube extends GoogleAppsScript.YouTube.YouTube {}

declare const YouTube: YouTube;
