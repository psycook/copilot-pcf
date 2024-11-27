import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as WebChat from 'botframework-webchat';

const DEFAULT_LOCALE = 'en-US';

export class CopilotComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    //----------------------------------------------------------------------------
    // Component properties
    //----------------------------------------------------------------------------

    // component internal properties
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;

    // component internal methods
    private _notifyOutputChanged: () => void;

    // component visual properties
    private _copilotContainer: HTMLDivElement | null = null;
    
    // bot framework web chat properties
    private _directLine: any;
    private _tokenEndpoint: string | undefined = undefined;
    private _copilotEndpoint: string | undefined = undefined;
    private _copilotToken: string | undefined = undefined;

    //component intenal properties
    private _isInitialised: boolean = false;
    private _debug: boolean = true;
    private _width: number;
    private _height: number;

    //----------------------------------------------------------------------------
    // Component lifecycle methods
    //----------------------------------------------------------------------------

    constructor()
    {
    }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        if(this._debug) console.log(`CopilotComponent:init() - start`);
        this._container = container;
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._context.mode.trackContainerResize(true);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        if(this._debug) console.log(`CopilotComponent:updateView() - start`);

        this.checkForCopilotUpdates(context);
        this.checkForUIUpdates(context);

        this._width = context.mode.allocatedWidth;
        this._height = context.mode.allocatedHeight;

    }

    public getOutputs(): IOutputs
    {
        if(this._debug) console.log(`CopilotComponent:getOutputs() - start`);

        return {};
    }

    public destroy(): void
    {
        if(this._debug) console.log(`CopilotComponent:destroy() - start`);
    }

    //----------------------------------------------------------------------------
    // Component custom methods
    //----------------------------------------------------------------------------

    private checkForCopilotUpdates(context: ComponentFramework.Context<IInputs>): void
    {
        if(this._debug) console.log(`CopilotComponent:checkForCopilotUpdates() - start`);

        this._context = context;

        if(this._context.parameters.tokenEndpoint.raw != null && this._context.parameters.tokenEndpoint.raw != this._tokenEndpoint)
        {
            this._tokenEndpoint = this._context.parameters.tokenEndpoint.raw;
            this._directLine = null;
            
            // end the conversation and remove the copilot container if it exists
            if(this._copilotContainer !== null)
            {
                this.endConversation();
                this._container.removeChild(this._copilotContainer);
                this._copilotContainer = null;
            }

            // create the copilot container
            this.getCopilotToken(new URL(this._tokenEndpoint)).then((token: string) => {
                this._directLine = WebChat.createDirectLine({ token });
                this.createCopilotContainer();
                this.renderCopilot();
                this.startConversation();
            });
        }
    }

    private checkForUIUpdates(context: ComponentFramework.Context<IInputs>): void
    {
        if(this._debug) console.log(`CopilotComponent:checkForUIUpdates() - start`);

        if(this._copilotContainer == null) return;

        var needsUpdate: boolean = false;
        
        if(context.mode.allocatedWidth != this._width && context.mode.allocatedHeight != this._height)
        {
            needsUpdate = true;
        }

        if(needsUpdate)
        {
            this.updateCopilotContainer();
        }
    }

    private createCopilotContainer(): void
    {
        if(this._debug) console.log(`CopilotComponent:createCopilotContainer() - start`);

        if(this._copilotContainer !== null) return;

        this._copilotContainer = document.createElement('div');
        this._copilotContainer.id = 'webchat';
        this._copilotContainer.role = 'main';
        this._copilotContainer.style.bottom = '0';
        this._copilotContainer.style.right = '0';
        this._copilotContainer.style.textAlign = 'left';
        this.updateCopilotContainer()
        this._container.appendChild(this._copilotContainer);
    }

    private updateCopilotContainer(): void
    {
        if(this._copilotContainer === null) return;

        //this._copilotContainer.style.fontSize = this._context.parameters.fontSize.raw || '18px';
        this._copilotContainer.style.height = `${this._context.mode.allocatedHeight}px`;
        this._copilotContainer.style.width = `${this._context.mode.allocatedWidth}px`;
    }

    private async getCopilotToken(tokenEndpointURL: URL): Promise<string>
    {
        if(this._debug) console.log(`CopilotComponent:getCopilotToken() - start`);

        try 
        {
            const response = await fetch(tokenEndpointURL.toString());
            if(!response.ok)
            {
                throw new Error(`CopilotComponent:getCopilotToken() - Error fetching token: ${response.statusText}`);
            }
            const {token} = await response.json();
            if(this._debug) console.log(`CopilotComponent:getCopilotToken() - token: ${token}`);
            return token;
        } 
        catch (error) 
        {
            console.error(`CopilotComponent:getCopilotToken() Error getting token direct line: ${JSON.stringify(error instanceof Error ? error.message : error)}`);
            return '';
        }
    }

    private renderCopilot(): void
    {
        const styleSet = WebChat.createStyleSet(
            {
                accent: "#000000",
                hideUploadButton: true,
                backgroundColor: "#F8F8F8",
                bubbleBorderColor: "#f08040",
                sendBoxButtonColor: "#000000",
                transcriptTerminatorFontSize: "24px",
                timestampColor: "#f08040",
                rootHeight: '100%',
                rootWidth: '100%',
        });
    
        const avatarOptions = {
            botAvatarImage: "https://github.com/psycook/images/blob/main/Contoso%20Bank%20Chatbot.png?raw=true",
            botAvatarInitials: "HB",
            userAvatarImage: "https://github.com/psycook/images/blob/main/Contoso%20Bank%20Customer%20New.png?raw=true",
            userAvatarInitials: "MS",
        };

        WebChat.renderWebChat(
            {
                directLine: this._directLine,
                userID: 'Simon',
                username: 'psycook',
                locale: DEFAULT_LOCALE,
                styleSet,
                styleOptions: avatarOptions
            },
            document.getElementById('webchat')
        );
    }

    private startConversation(): void
    {
        if(this._debug) console.log(`CopilotComponent:startConversation() - start`);

        if(this._directLine == null) return;    

        this._directLine.postActivity({
            from: { id: 'Simon', name: 'psycook' },
            name: 'requestWelcomeDialog',
            type: 'event',
            value: ''
        }).subscribe(
            (id:string) => console.log(`Posted activity, assigned ID ${id}`),
            (error:string) => console.log(`Error posting activity ${error}`)
        );
    }

    private endConversation(): void
    {
        if(this._debug) console.log(`CopilotComponent:endConversation() - start`);
        if(this._directLine == null) return;
        this._directLine.end();
    }
}
