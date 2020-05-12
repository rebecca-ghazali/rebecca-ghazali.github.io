import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/homeTab/index.html")
@PreventIframe("/homeTab/config.html")
@PreventIframe("/homeTab/remove.html")
export class HomeTab {
}
