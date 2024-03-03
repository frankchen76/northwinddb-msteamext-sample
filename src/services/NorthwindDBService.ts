import { HttpClientService } from "./HttpClientService"

export interface IProductQuery {
    productName: string | null;
    categoryName: string | null;
    inventoryStatus: string | null;
    supplierCity: string | null;
    stockLevel: string | null;
}

export class NorthwindDBService extends HttpClientService {
    constructor(credential, scopes) {
        super(credential, scopes);
    }

    public async getProducts(query: IProductQuery): Promise<any> {
        const url = `${process.env.NORTHWINDDBAPI_ENDPOINT}/api/products`;
        const body = {
            "productQuery": query
        };
        const response = await this.post(url, body);
        return response;
    }

}