import * as React from 'react';
import { AadHttpClient, AadHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import useGraphClient from './useGraphClient';

export interface IExecuteBatchRequest {
  method: string;
  url: string;
  id: string | number;
  headers?: HeadersInit;
  body?: string;
}

export interface IExecuteBatchResponse {
  id: string | number;
  status: string;
  body: string;
}

export interface IExecuteBatchReturn {
  executeBatch: (method: string, requests: IExecuteBatchRequest[]) => Promise<IExecuteBatchResponse[]>;
}

const useBatch = (context: WebPartContext): IExecuteBatchReturn => {
  const { client } = useGraphClient(context);

  const executeBatch = React.useCallback(
    async (
      method: string,
      requests: IExecuteBatchRequest[]
    ): Promise<IExecuteBatchResponse[]> => {
      const batchBody = {
        requests: requests.map((item: IExecuteBatchRequest) => ({
          id: item.id,
          method,
          url: item.url,
          headers: item.headers ?? {},
          body: item.body ?? {}
        }))
      };

      try {
        if (client) {
          const res: AadHttpClientResponse = await client.post(
            `https://graph.microsoft.com/v1.0/$batch`,
            AadHttpClient.configurations.v1,
            {
              headers: {
                Accept: 'application/json',
                'Content-Type': 'application/json'
              },
              body: JSON.stringify(batchBody)
            }
          );

          const json = await res.json();

          const responses: IExecuteBatchResponse[] = json.responses.map(
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (val: any) => ({
              id: val.id,
              status: val.status,
              body: val.body
            })
          );

          return responses;
        } else {
          return [];
        }
      } catch (e) {
        console.error(e);

        return [];
      }
    },
    [client]
  );

  return { executeBatch };
};

export default useBatch;
