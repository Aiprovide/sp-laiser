/// <reference types="sharepoint" />
import "Promise";
import { Bindable } from "laiser";
export declare class PeopleEditor extends Bindable<SP.User> {
    private _SPClientContext;
    private _SPWeb;
    private _ErrorMsg;
    constructor(id: string, parent: any, je: boolean, spClientContext: SP.ClientContext, spWeb: SP.Web);
    get_Value(): Promise<SP.User>;
    ToString(): string;
    MessageError(message: string): void;
    set_Text(value: SP.User): void;
    set_Css(value: SP.User): void;
    set_TextCss(value: SP.User): void;
}
export declare class InternalPeoplePicker extends Bindable<SP.User[]> {
    private _SPClientContext;
    private _SPWeb;
    private _ErrorMsg;
    constructor(id: string, parent: any, je: boolean, spClientContext: SP.ClientContext, spWeb: SP.Web, Multiple?: boolean, User?: boolean, SPGroup?: boolean, SecGroup?: boolean, DL?: boolean, width?: number);
    get_Value(): Promise<SP.User[]>;
    set_Value(users: SP.User[]): void;
    set_ValueByLoginName(loginNames: string[]): void;
    MessageError(message: string): void;
    set_Text(value: SP.User[]): void;
    set_Css(value: SP.User[]): void;
    set_TextCss(value: SP.User[]): void;
}
export declare class ExternalPeoplePicker extends Bindable<SP.User[]> {
    private _SPClientContext;
    private _SPWeb;
    private _ErrorMsg;
    private _ClientID;
    private _PeoplePicker;
    constructor(id: string, parent: any, je: boolean, spClientContext: SP.ClientContext, spWeb: SP.Web, hostUrl: string, appUrl: string, language: string, Multiple?: boolean);
    get_Value(): Promise<SP.User[]>;
    set_Value(users: SP.User[]): void;
    set_ValueByLoginName(loginNames: string[]): void;
    MessageError(message: string): void;
    set_Text(value: SP.User[]): void;
    set_Css(value: SP.User[]): void;
    set_TextCss(value: SP.User[]): void;
}
