export interface ExampleType {
    id: number;
    name: string;
    description?: string;
}

export type User = {
    id: string;
    username: string;
    email: string;
};

export interface Product {
    id: string;
    title: string;
    price: number;
    inStock: boolean;
}