import { TeamsActivityHandler, CardFactory, TurnContext, MessagingExtensionQuery, MessagingExtensionResponse } from "botbuilder";
// import {
//     MessageExtensionTokenResponse,
//     handleMessageExtensionQueryWithSSO,
//     OnBehalfOfCredentialAuthConfig,
//     OnBehalfOfUserCredential,
// } from "@microsoft/teamsfx";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
// import { AccessToken, GetTokenOptions, TokenCredential } from "@azure/core-auth";
import "isomorphic-fetch";
import { IProductQuery, NorthwindDBService } from "../services/NorthwindDBService";
import cardHandler from "../adaptiveCards/cardHandler";
import { log } from '../log';
import { Query } from "@microsoft/teams-ai";

export interface ITokenCredential {
    getToken(scopes: string | string[]): Promise<string>;
}
export class TeamAICredential implements ITokenCredential {
    constructor(private readonly token: string) {
    }
    async getToken(scopes: string | string[]): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            setTimeout(() => {
                resolve(this.token);
            }, 100);

        });
    }

}
export class TeamsExtService {
    private queryCount = 0;

    public async searchProductsFromCopilot(context: TurnContext, query: Query<Record<string, string>>, credential: ITokenCredential): Promise<MessagingExtensionResponse> {
        log('query', query);

        // let productName, categoryName, inventoryStatus, supplierCity, stockLevel;

        // if (query.parameters.length === 1 && query.parameters.productName!=null) {
        //     [productName, categoryName, inventoryStatus, supplierCity, stockLevel] = (query.parameters.productName.split(','));
        // } else {
        //     productName = this.cleanupParam(query.parameters.find((element) => element.name === "productName")?.value);
        //     categoryName = this.cleanupParam(query.parameters.find((element) => element.name === "categoryName")?.value);
        //     inventoryStatus = this.cleanupParam(query.parameters.find((element) => element.name === "inventoryStatus")?.value);
        //     supplierCity = this.cleanupParam(query.parameters.find((element) => element.name === "supplierCity")?.value);
        //     stockLevel = this.cleanupParam(query.parameters.find((element) => element.name === "stockQuery")?.value);
        // }


        const productQuery: IProductQuery = {
            productName: query.parameters.productName ?? null,
            categoryName: query.parameters.categoryName ?? null,
            inventoryStatus: query.parameters.inventoryStatus ?? null,
            supplierCity: query.parameters.supplierCity ?? null,
            stockLevel: query.parameters.stockQuery ?? null,
        };
        const response = await this.searchProduct(context, productQuery, credential);
        return response;

    }
    // public async searchProductsFromExt(context: TurnContext, query: MessagingExtensionQuery, credential: TokenCredential): Promise<MessagingExtensionResponse> {

    //     let productName, categoryName, inventoryStatus, supplierCity, stockLevel;

    //     // For now we have the ability to pass parameters comma separated for testing until the UI supports it.
    //     // So try to unpack the parameters but when issued from Copilot or the multi-param UI they will come
    //     // in the parameters array.
    //     if (query.parameters.length === 1 && query.parameters[0]?.name === "productName") {
    //         [productName, categoryName, inventoryStatus, supplierCity, stockLevel] = (query.parameters[0]?.value.split(','));
    //     } else {
    //         productName = this.cleanupParam(query.parameters.find((element) => element.name === "productName")?.value);
    //         categoryName = this.cleanupParam(query.parameters.find((element) => element.name === "categoryName")?.value);
    //         inventoryStatus = this.cleanupParam(query.parameters.find((element) => element.name === "inventoryStatus")?.value);
    //         supplierCity = this.cleanupParam(query.parameters.find((element) => element.name === "supplierCity")?.value);
    //         stockLevel = this.cleanupParam(query.parameters.find((element) => element.name === "stockQuery")?.value);
    //     }
    //     log(`🔎 Query #${++this.queryCount}:\nproductName=${productName}, categoryName=${categoryName}, inventoryStatus=${inventoryStatus}, supplierCity=${supplierCity}, stockLevel=${stockLevel}`);

    //     const productQuery: IProductQuery = {
    //         productName: productName,
    //         categoryName: categoryName,
    //         inventoryStatus: inventoryStatus,
    //         supplierCity: supplierCity,
    //         stockLevel: stockLevel,
    //     };
    //     return await this.searchProduct(context, productQuery, credential);
    // }
    private async searchProduct(context: TurnContext, productQuery: IProductQuery, credential: ITokenCredential): Promise<MessagingExtensionResponse> {
        const service = new NorthwindDBService(credential, ["api://ab4e8ed8-c4d7-4de9-a352-d23da8651cf9/.default"]);

        const products = await service.getProducts(productQuery);

        log(`Found ${products.length} products in the Northwind database`);
        const attachments = [];
        products.forEach((product) => {
            const preview = CardFactory.heroCard(product.ProductName,
                `Supplied by ${product.SupplierName} of ${product.SupplierCity}<br />${product.UnitsInStock} in stock`,
                [product.ImageUrl]);

            const resultCard = cardHandler.getEditCard(product);
            const attachment = { ...resultCard, preview };
            attachments.push(attachment);
        });
        return {
            composeExtension: {
                type: "result",
                attachmentLayout: "list",
                attachments: attachments,
            },
        };

    }
    private cleanupParam(value: string): string {

        if (!value) {
            return "";
        } else {
            let result = value.trim();
            result = result.split(',')[0];          // Remove extra data
            result = result.replace("*", "");       // Remove wildcard characters from Copilot
            return result;
        }
    }


    public async handleTeamsMessagingExtensionSelectItem(
        context: TurnContext,
        obj: any
    ): Promise<any> {
        return {
            composeExtension: {
                type: "result",
                attachmentLayout: "list",
                attachments: [CardFactory.heroCard(obj.name, obj.description)],
            },
        };
    }
}