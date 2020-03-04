import * as React from 'react';
import styles from './DownloadIcsExample.module.scss';
import * as strings from 'DownloadIcsExampleWebPartStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDownloadIcsExampleProps, IIcsEvent } from './DownloadIcsExample.Types';
import { DefaultButton } from 'office-ui-fabric-react';

/**
 * Enables a event to be downloaded as a ICS file to be imported to the users calendar
 * Inspired by: https://github.com/josephj/react-icalendar-link/tree/master/src
 */
export default class DownloadIcsExample extends React.Component<IDownloadIcsExampleProps, {}> {
  public render(): React.ReactElement<IDownloadIcsExampleProps> {
    return (
      <div className={ styles.downloadIcsExample }>
        <DefaultButton
          text={strings.AddEventToCalendarLabel}
          onClick={this.handleDownloadIcsClick}
          iconProps={{ iconName: 'AddEvent' }}
        />
      </div>
    );
  }

  /** Event: Download ICS file */
  private handleDownloadIcsClick = () => {
    // TODO: Replace this example with event from in example SharePoint list
    const exampleEvent: IIcsEvent = {
      title: "National Championship Downhill Race",
      location: "Steep Hills Alpine Resort",
      startTime: new Date("04.02.2020 10:00"),
      endTime: new Date("04.02.2020 14:00"),
      description: "For more information, please visit: https://www.google.com",
      url: "https://wwww.google.com"
    };

    const icsEvent: string = this.mapEventToIcs(exampleEvent);
    const blob: object = new Blob([icsEvent], {
        type: "text/calendar;charset=utf-8"
    });

    // IE
    if (this.isIE()) {
        window.navigator.msSaveOrOpenBlob(blob, "calendar.ics");
        return;
    }

    // Safari
    if (this.isIOSSafari()) {
        window.open(encodeURI(`data:text/calendar;charset=utf8,${icsEvent}`), "_blank");
        return;
    }

    // Desktop
    this.downloadBlob(blob, "calendar.ics");    
  }

  /** Create BLOB object and download in the browser */
  private downloadBlob(blob: object, filename: string): void {
    const element: HTMLAnchorElement = document.createElement("a");
    const url = window.URL.createObjectURL(blob);
    element.href = url;
    element.setAttribute("download", filename);
    element.click();
    window.URL.revokeObjectURL(url);
    element.remove();
}

  private mapEventToIcs = (event: IIcsEvent): string => {
    const icsEvent: string = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "BEGIN:VEVENT",
        "URL:" + event.url,
        "DTSTART:" + this.formatDate(event.startTime),
        "DTEND:" + this.formatDate(event.endTime),
        "SUMMARY:" + event.title,
        "DESCRIPTION:" + event.description,
        "LOCATION:" + event.location,
        "END:VEVENT",
        "END:VCALENDAR"
    ].join("\n");

    return icsEvent;
  }

  private formatDate (dateTime: Date): string {
      return [
          dateTime.getUTCFullYear(),
          this.pad(dateTime.getUTCMonth() + 1),
          this.pad(dateTime.getUTCDate()),
          "T",
          this.pad(dateTime.getUTCHours()),
          this.pad(dateTime.getUTCMinutes()) + "00Z"
      ].join("");
  }

  private pad(num: number): string {
      if (num < 10) {
      return `0${num}`;
      }
      return `${num}`;
  }

  private isIE(): boolean {
    return !!(
    typeof window !== "undefined" &&
    window.navigator.msSaveOrOpenBlob &&
    window.Blob
    );
  }

  private isIOSSafari(): boolean {
      const ua = window.navigator.userAgent;
      const iOS = !!ua.match(/iPad/i) || !!ua.match(/iPhone/i);
      const webkit = !!ua.match(/WebKit/i);

      return iOS && webkit && !ua.match(/CriOS/i);
  }
}
