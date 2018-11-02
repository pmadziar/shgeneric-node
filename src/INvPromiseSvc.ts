import { Promise } from "es6-promise";

export interface INvPromiseSvc<T> {
	GetAsync: () => Promise<INvPromiseSvc<T>>;
	ClientContext: SP.ClientContext;
	Site: INvPromiseSvc<SP.Site>;
	Web: INvPromiseSvc<SP.Web>;
	List: INvPromiseSvc<SP.List>;
	Target: T;
}
