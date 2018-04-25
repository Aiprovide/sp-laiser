import "Promise";
//import $ from "jquery";
import {Common} from "sp-com";
import {Bindable, Label, Com} from "laiser";

//
// PeopleEditor
//
export class PeopleEditor extends Bindable<SP.User>
{
    // サイトコレクションとトップレベルサイトのコンテキスト
    private _SPClientContext: SP.ClientContext = null;
    private _SPWeb: SP.Web = null;
    private _ErrorMsg: Label;

    constructor(id: string = null, parent: any = null, je: boolean = true, spClientContext:SP.ClientContext, spWeb: SP.Web)
    {
        super(id, parent, je);
        this._SPClientContext = spClientContext;
        this._SPWeb = spWeb;
        if (id !== null)
        {
            let coreid = "div[id $= '" + id + "_upLevelDiv'] div#divEntityData";
            this._ErrorMsg = new Label("span[id $= '" + id + "_errorLabel']", this, true, "span");
        }
    }

    async get_Value(): Promise<SP.User>
    {
        this._JElem = (this.ID !== null) ? $(Com.ID(this.ID)) : null;
        if (this._JElem !== null && typeof this._JElem !== "undefined" && this._JElem.length > 0)
        {
            let key: string = this._JElem.attr("key");
            if (key === null || typeof key === "undefined" || key.length === 0)
            {
                return null;
            }
            let user: SP.User = await Common.Get_SPGivenUser(this._SPClientContext, key, this._SPWeb);
            return user;
        }
        else
        {
            return null;
        }
    }

    ToString(): string
    {
        let strOut = "";
        let strCss = (this._CssClass.length > 0) ? " class='" + this._CssClass.join(" ") + "'" : "";
        let strAttribute = "";
        this._Attribute.forEach((avalue: { key: string; value: string }, index: number, array: { key: string; value: string }[]): void =>
        {
            strAttribute += " " + avalue.key + "='" + avalue.value + "'";
        });

        if (this._Id !== null && this._Id.length > 0 && typeof this._Id !== "undefined")
        {
            strOut = "<SharePoint:PeopleEditor runat='server' id='" + this._Id + "' " + strCss + strAttribute + " />";
        }
        else
        {
            strOut = "<SharePoint:PeopleEditor runat='server' " + strCss + strAttribute + " />";
        }
        return strOut;
    }

    //
    // エラーメッセージを表示する
    //
    MessageError(message: string): void
    {
        this._ErrorMsg.Text = message;
    }

    // 表示内容設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_Text(value: SP.User): void
    {
        if (this._BmTextCss === null)
        {
            if (this._BmText !== null)
            {
                this._BmText(value);
            }
            else
            {
                this.set_Value(value);
            }
        }
    }
    // CSS設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_Css(value: SP.User): void
    {
        if (this._BmCss !== null)
        {
            this._BmCss(value);
        }
    }
    // 表示内容とCSSの同時設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_TextCss(value: SP.User): void
    {
        if (this._BmTextCss !== null)
        {
            this._BmTextCss(value);
        }
    }
}

//
// SharePoint Hosted Appで使用する
// PeoplePicker
//
export class InternalPeoplePicker extends Bindable<SP.User[]>
{
    // サイトコレクションとトップレベルサイトのコンテキスト
    private _SPClientContext: SP.ClientContext = null;
    private _SPWeb: SP.Web = null;
    private _ErrorMsg: Label;

    constructor(id: string = null, parent: any = null, je: boolean = true, spClientContext:SP.ClientContext, spWeb: SP.Web, Multiple: boolean = true, User: boolean = true, SPGroup: boolean = true, SecGroup: boolean = true, DL: boolean = true, width: number = 200)
    {
        super(id, parent, je);
        this._SPClientContext = spClientContext;
        this._SPWeb = spWeb;
        
        if (id !== null)
        {
            // Create a schema to store picker properties, and set the properties.
            let schema = {};
            let accouttype: string = "";
            accouttype += (User) ? "User" : "";
            accouttype += (DL) ? ",DL" : "";
            accouttype += (SecGroup) ? ",SecGroup" : "";
            accouttype += (SPGroup) ? ",SPGroup" : "";
            schema['PrincipalAccountType'] = accouttype;
            schema['SearchPrincipalSource'] = 15;
            schema['ResolvePrincipalSource'] = 15;
            schema['AllowMultipleValues'] = Multiple;
            schema['MaximumEntitySuggestions'] = 50;
            schema['Width'] = width.toString() + 'px';

            // Render and initialize the picker. 
            window.SPClientPeoplePicker_InitStandaloneControlWrapper(id, null, schema);

            this._ErrorMsg = new Label(id + "_errorLabel", this, true);
        }
    }

    async get_Value(): Promise<SP.User[]>
    {
        let coreid: string = this.ID + "_TopSpan";
        let peoplePicker: any = window.SPClientPeoplePicker.SPClientPeoplePickerDict[coreid];
        if (peoplePicker === null || typeof peoplePicker === "undefined")
        {
            return;
        }
        let userInfos: any[] = peoplePicker.GetAllUserInfo();
        let loginNames: string[] = Array<string>();
        userInfos.forEach((value: any, index: number, array: any[]): void =>
        {
            let loginName: string = value.Key;
            loginNames.push(loginName);
        });

        if (loginNames.length > 0)
        {
            let users: SP.User[] = await Common.Get_SPGivenUsers(this._SPClientContext, loginNames, this._SPWeb);
            return users;
        }
        else
        {
            return null;
        }
    }

    set_Value(users: SP.User[]): void
    {
        let coreid: string = this.ID + "_TopSpan";
        let peoplePicker: any = window.SPClientPeoplePicker.SPClientPeoplePickerDict[coreid];
        if (peoplePicker === null || typeof peoplePicker === "undefined")
        {
            return;
        }

        users.forEach((user: SP.User, index: number, array: SP.User[]): void =>
        {
            peoplePicker.AddUnresolvedUser({ Key: user.get_loginName() }, true);
        });

        return;
    }

    set_ValueByLoginName(loginNames: string[]): void
    {
        let coreid: string = this.ID + "_TopSpan";
        let peoplePicker: any = window.SPClientPeoplePicker.SPClientPeoplePickerDict[coreid];
        if (peoplePicker === null || typeof peoplePicker === "undefined")
        {
            return;
        }

        loginNames.forEach((loginName: string, index: number, array: string[]): void =>
        {
            peoplePicker.AddUnresolvedUser({ Key: loginName }, true);
        });

        return;
    }

    //
    // エラーメッセージを表示する
    //
    MessageError(message: string): void
    {
        this._ErrorMsg.Text = message;
    }

    // 表示内容設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_Text(value: SP.User[]): void
    {
        if (this._BmTextCss === null)
        {
            if (this._BmText !== null)
            {
                this._BmText(value);
            }
            else
            {
                this.set_Value(value);
            }
        }
    }
    // CSS設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_Css(value: SP.User[]): void
    {
        if (this._BmCss !== null)
        {
            this._BmCss(value);
        }
    }
    // 表示内容とCSSの同時設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_TextCss(value: SP.User[]): void
    {
        if (this._BmTextCss !== null)
        {
            this._BmTextCss(value);
        }
    }
}

//
// Provider Hosted Appで使用する
// PeoplePicker
//
export class ExternalPeoplePicker extends Bindable<SP.User[]>
{
    // サイトコレクションとトップレベルサイトのコンテキスト
    private _SPClientContext: SP.ClientContext = null;
    private _SPWeb: SP.Web = null;
    private _ErrorMsg: Label;
    private _ClientID: string;
    private _PeoplePicker: any;

    constructor(id: string = null, parent: any = null, je: boolean = true, spClientContext:SP.ClientContext, spWeb: SP.Web, hostUrl: string, appUrl: string, language: string, Multiple: boolean = true)
    {
        super(id, parent, je);
        this._SPClientContext = spClientContext;
        this._SPWeb = spWeb;

        if (id !== null)
        {
            this._ErrorMsg = new Label(id + "_errorLabel", this, true);

            //let spanAdministrators: JQuery = $(Com.ID(id + "_spanAdministrators"));
            //let inputAdministrators: JQuery = $(Com.ID(id + "_inputAdministrators"));
            //let divAdministratorsSearch: JQuery = $(Com.ID(id + "_divAdministratorsSearch"));
            //let hdnAdministrators: JQuery = $(Com.ID(id + "_hdnAdministrators"));

            //this._PeoplePicker = new CAMControl.PeoplePicker(context, hostUrl, spanAdministrators, inputAdministrators, divAdministratorsSearch, hdnAdministrators);
            //// required to pass the variable name here!
            //this._PeoplePicker.InstanceName = "peoplePicker";
            //this._PeoplePicker.Language = language;
            //this._PeoplePicker.AllowDuplicates = Multiple;
            //// Hookup everything
            //this._PeoplePicker.Initialize();

            //let loginConfig: Office.Controls.LoginConfig =
            //    {
            //        instance: 'https://login.microsoftonline.com/',
            //        clientId: this._ClientID, //Please replace with your clientID
            //        redirectUri: window.location.href,
            //        postLogoutRedirectUri: window.location,
            //        cacheLocation: 'localStorage' // enable this for IE, as sessionStorage does not work for localhost.
            //    };
            //let authContext: any = new window.AuthenticationContext(loginConfig);
            //let isCallback = authContext.isCallback(window.location.hash);
            //authContext.handleWindowCallback();
            //let user: any = authContext.getCachedUser();

            //var aadDataProvider: Office.Controls.PeopleAadDataProvider = new Office.Controls.PeopleAadDataProvider(authContext);

            //let poeplePickerOptions: Office.Controls.PeoplePickerOptions =
            //    {
            //        allowMultipleSelections: Multiple,
            //        startSearchCharLength: 1, 
            //        inputHint: "検索..."
            //    };
            //this._PeoplePicker = new Office.Controls.PeoplePicker(document.getElementById(id), aadDataProvider, poeplePickerOptions);


            let runtimeOptions: Office.Controls.RuntimeOptions =
                {
                    sharePointHostUrl: hostUrl,
                    appWebUrl: appUrl
                };

            Office.Controls.Runtime.initialize(runtimeOptions);
            Office.Controls.Runtime.renderAll();

            let poeplePickerOptions: Office.Controls.PeoplePickerOptions =
                {
                    allowMultipleSelections: Multiple,
                    placeholder: "Enter names or email addresses..."
                };

            // Render PoeplePicker. 
            this._PeoplePicker = new Office.Controls.PeoplePicker(document.getElementById(id), poeplePickerOptions);
        }
    }

    //createLoginProvider(): Office.Controls.ImplicitGrantLogin
    //{
    //    let loginConfig: Office.Controls.LoginConfig =
    //        {
    //            instance: 'https://login.microsoftonline.com/',
    //            clientId: this._ClientID, //Please replace with your clientID
    //            redirectUri: window.location.href,
    //            postLogoutRedirectUri: window.location,
    //            cacheLocation: 'localStorage' // enable this for IE, as sessionStorage does not work for localhost.
    //        };

    //    var loginProvider = new Office.Controls.ImplicitGrantLogin(loginConfig);
    //    return loginProvider;
    //}

    async get_Value(): Promise<SP.User[]>
    {
        if (this._PeoplePicker === null)
        {
            return null;
        }

        let pickedUsers: Office.Controls.PeoplePickerRecord[] = this._PeoplePicker.selectedItems;
        //let pickedUsers: Office.Controls.PeoplePickerRecord[] = this._PeoplePicker.getAddedPeople();

        let loginNames: string[] = Array<string>();
        pickedUsers.forEach((value: Office.Controls.PeoplePickerRecord, index: number, array: Office.Controls.PeoplePickerRecord[]): void =>
        {
            let loginName: string = value.loginName;
            loginNames.push(loginName);
        });

        if (loginNames.length > 0)
        {
            let users: SP.User[] = await Common.Get_SPGivenUsers(this._SPClientContext, loginNames, this._SPWeb);
            return users;
        }
        else
        {
            return null;
        }
    }

    set_Value(users: SP.User[]): void
    {
        if (this._PeoplePicker === null)
        {
            return;
        }

        this._PeoplePicker.reset();

        users.forEach((user: SP.User, index: number, array: SP.User[]): void =>
        {
            let record: Office.Controls.PeoplePickerRecord = new Office.Controls.PeoplePickerRecord();
            record.displayName = user.get_title();
            record.loginName = user.get_loginName();
            record.email = user.get_email();
            this._PeoplePicker.add(record, true);
        });

        return;
    }

    set_ValueByLoginName(loginNames: string[]): void
    {
        if (this._PeoplePicker === null)
        {
            return;
        }

        this._PeoplePicker.reset();

        loginNames.forEach((loginName: string, index: number, array: string[]): void =>
        {
            let record: Office.Controls.PeoplePickerRecord = new Office.Controls.PeoplePickerRecord();
            record.displayName = "";
            record.loginName = loginName;
            record.email = "";
            this._PeoplePicker.add(record, true);
        });

        return;
    }

    //
    // エラーメッセージを表示する
    //
    MessageError(message: string): void
    {
        this._ErrorMsg.Text = message;
    }

    // 表示内容設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_Text(value: SP.User[]): void
    {
        if (this._BmTextCss === null)
        {
            if (this._BmText !== null)
            {
                this._BmText(value);
            }
            else
            {
                this.set_Value(value);
            }
        }
    }
    // CSS設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_Css(value: SP.User[]): void
    {
        if (this._BmCss !== null)
        {
            this._BmCss(value);
        }
    }
    // 表示内容とCSSの同時設定用
    // バインド先のデータが変更されていた場合は、データから呼ばれる
    set_TextCss(value: SP.User[]): void
    {
        if (this._BmTextCss !== null)
        {
            this._BmTextCss(value);
        }
    }
}
