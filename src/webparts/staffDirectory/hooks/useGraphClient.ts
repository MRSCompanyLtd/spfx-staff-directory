import * as React from 'react';
import { AadHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IUseGraphClient {
    client: AadHttpClient | undefined;
}

const useGraphClient = (context: WebPartContext): IUseGraphClient => {
  const [client, setClient] = React.useState<AadHttpClient>();

  const getClient = React.useCallback(async () => {
    const c: unknown = await context.aadHttpClientFactory.getClient(
      'https://graph.microsoft.com'
    ).catch(e => {
        console.error(e);

        throw Error(e);
    });

    setClient(c as AadHttpClient);
  }, [context]);

  React.useEffect(() => {
    const get = async (): Promise<void> => {
        await getClient();
    }
    
    get().catch(e => {
        console.error(e);

        throw Error(e);
    });
  }, [getClient]);

  return { client };
};

export default useGraphClient;
