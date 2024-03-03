import axios, { type AxiosRequestConfig } from 'axios'
import * as https from 'https';
import { ITokenCredential } from '../sso/TeamsExtService';
// import { OnBehalfOfUserCredential } from "@microsoft/teamsfx";
// import { TokenCredential } from "@azure/core-auth";
export enum HttpClientAuthType {
    Bearer = 'Authorization',
    ApiKey = 'api-key'
}
export class HttpClientService {
    // constructor(private readonly accessToken: string, private readonly authType: HttpClientAuthType = HttpClientAuthType.Bearer) {

    // }
    constructor(private readonly credential: ITokenCredential, private scopes: string[]) {

    }

    protected async getAxiosConfig(): Promise<AxiosRequestConfig> {
        const accesstoken = await this.credential.getToken(this.scopes);
        console.log("getAxiosConfig", accesstoken);
        const standardHeaders = {
            'Content-Type': 'application/json',
            cache: 'no-store'
        }
        //const headers = this.authType === HttpClientAuthType.Bearer ? { ...standardHeaders, Authorization: `Bearer ${this.accessToken}` } : { ...standardHeaders, 'x-api-key': `${this.accessToken}` }
        const headers = { ...standardHeaders, Authorization: `Bearer ${accesstoken}` }
        const config: AxiosRequestConfig = {
            headers: headers,
            httpsAgent: new https.Agent({
                rejectUnauthorized: false
            })
        }
        return config
    }

    public async get(url: string): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.get(url, config)
        return response.data
    }

    public async post(url: string, body: any, isAdd?: boolean): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.post(url, body, config)
        return response.data
    }

    public async patch(url: string, body: any, isAdd?: boolean): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.patch(url, body, config)
        return response.data
    }

    public async delete(url: string): Promise<any> {
        const config = await this.getAxiosConfig();
        const response = await axios.delete(url, config)
        return response.data
    }
}
