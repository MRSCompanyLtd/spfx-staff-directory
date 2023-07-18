import * as React from 'react';
import useGraphClient from './useGraphClient';
import {
  AadHttpClient,
  AadHttpClientResponse,
  IAadHttpClientConfiguration
} from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPerson } from '../interfaces/IPerson';

interface IUseSearch {
  searchByText: (str: string, filterDepartment?: string) => Promise<void>;
  getNextPage: () => Promise<void>;
  total: number;
  loading: boolean;
  nextPage: string;
  results: IPerson[];
}

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

const GRAPH_URL = 'https://graph.microsoft.com/v1.0';
const OPTIONS: IAadHttpClientConfiguration = {
  headers: {
    ConsistencyLevel: 'Eventual',
    Accept: 'application/json',
    'Content-Type': 'application/json'
  }
};
const SELECT = [
  'id',
  'displayName',
  'department',
  'jobTitle',
  'businessPhones',
  'mail',
  'userPrincipalName'
];

const useSearch = (
  context: WebPartContext,
  group: string,
  pageSize: number
): IUseSearch => {
  const [loading, setLoading] = React.useState<boolean>(false);
  const [nextPage, setNextPage] = React.useState<string>('');
  const [results, setResults] = React.useState<IPerson[]>([]);
  const [total, setTotal] = React.useState<number>(0);

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
        if (!client) return [];

        const res: AadHttpClientResponse = await client.post(
          `${GRAPH_URL}/$batch`,
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
      } catch (e) {
        console.error(e);

        return [];
      }
    },
    [client]
  );

  const fetchUserImages = React.useCallback(
    async (people: IPerson[]): Promise<IPerson[]> => {
      const requests: IExecuteBatchRequest[] = people.map((p) => ({
        url: `/users/${p.id}/photo/$value`,
        method: 'GET',
        id: p.id
      }));

      const responses = await executeBatch('GET', requests);

      const lookup: Record<string, string> = {};

      responses.forEach((r: IExecuteBatchResponse) => {
        lookup[r.id] = r.body;
      });

      const updatedPeople = people.map((p) => ({
        ...p,
        picture: Object.prototype.hasOwnProperty.call(lookup, p.id)
          ? lookup[p.id]
          : undefined
      }));

      return updatedPeople;
    },
    [executeBatch]
  );

  const searchByText = React.useCallback(
    async (str: string, filterDepartment = '') => {
      try {
        if (!client) return;

        setLoading(true);

        let url: string = `${GRAPH_URL}/`;

        if (group !== '') {
          url += `groups/${group}/members?`;
        } else {
          url += `users?`;
        }

        if (str && str !== '') {
          url += `$search="displayName:${str}" OR "department:${str}" OR "jobTitle:${str}"&`;
        }

        url += `$top=${pageSize}&$select=${SELECT.join(',')}&$count=true&`;

        if (filterDepartment !== '') {
          url += `$filter=department eq '${filterDepartment}'`;
        }

        const res = await client
          .get(`${url}`, AadHttpClient.configurations.v1, OPTIONS)
          .catch((e) => {
            console.error(e);

            throw Error(e);
          });

        if (res.ok) {
          const values = await res.json();
          const people =
            values.value.length > 0 ? await fetchUserImages(values.value) : [];

          const nextLink = values['@odata.nextLink'];
          const count = values['@odata.count'];

          setTotal(count);
          setNextPage(nextLink);
          setResults(people);

          setLoading(false);
        }
      } catch (e) {
        console.error('Error searching by text:', e);

        setLoading(false);
      }
    },
    [client, executeBatch, group, pageSize]
  );

  const getNextPage = React.useCallback(async () => {
    try {
      if (!client || !nextPage) return;

      setLoading(true);

      const res: AadHttpClientResponse = await client.get(
        `${nextPage}`,
        AadHttpClient.configurations.v1,
        OPTIONS
      );

      if (res.ok) {
        const values = await res.json();
        const people =
          values.value.length > 0 ? await fetchUserImages(values.value) : [];
        const nextLink = values['@odata.nextLink'];

        setNextPage(nextLink);
        setResults(people);
        setLoading(false);
      }
    } catch (e) {
      console.error('Error getting next page:', e);

      setLoading(false);
    }
  }, [client, nextPage]);

  React.useEffect(() => {
    if (client) {
      searchByText('').catch((e) => console.error(e));
    }
  }, [client]);

  return {
    searchByText,
    getNextPage,
    total,
    loading,
    nextPage,
    results
  };
};

export default useSearch;
