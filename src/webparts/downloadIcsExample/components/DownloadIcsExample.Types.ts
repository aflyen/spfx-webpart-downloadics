export interface IDownloadIcsExampleProps {
    description: string;
}

export interface IIcsEvent {
    title: string;
    startTime: Date;
    endTime: Date;
    description: string;
    url: string;
    location: string;
  }