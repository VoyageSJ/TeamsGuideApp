import { PreventIframe } from "express-msteams-host";

@PreventIframe("/YouTubePlayerTab/selector.html")

export class VideoSelectorTaskModule { }
