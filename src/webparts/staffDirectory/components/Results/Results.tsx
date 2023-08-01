import * as React from 'react';
import { IResultsProps } from './IResultsProps';
import styles from './Results.module.scss';
import { Spinner, SpinnerSize, Text } from 'office-ui-fabric-react';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { PersonCardInteraction, PersonViewType } from '@microsoft/mgt-spfx';

const Results: React.FC<IResultsProps> = ({ results, loading }) => {
  const getPhoto = (data: string): string => {
    const byteCharacters = atob(data);
    const byteArrays = [];

    for (let offset = 0; offset < byteCharacters.length; offset += 512) {
      const slice = byteCharacters.slice(offset, offset + 512);
      const byteNumbers = new Array(slice.length);

      for (let i = 0; i < slice.length; i++) {
        byteNumbers[i] = slice.charCodeAt(i);
      }

      const byteArray = new Uint8Array(byteNumbers);
      byteArrays.push(byteArray);
    }

    const blob: Blob = new Blob(byteArrays, { type: 'image/jpeg' });

    const url = window.URL || window.webkitURL;
    const img = url.createObjectURL(blob);

    return img;
  };

  return (
    <div className={styles.results}>
      {loading ? (
        <div
          style={{ width: '100%', display: 'flex', justifyContent: 'center' }}
        >
          <Spinner size={SpinnerSize.large} />
        </div>
      ) : results.length > 0 ? (
        <div className={styles.searchResults}>
          {results.map((item) => (
            <div className={styles.result} key={item.id}>
              <Person
                userId={item.id}
                personDetails={{
                  ...item,
                  personImage: getPhoto(item.picture ?? '')
                }}
                personCardInteraction={PersonCardInteraction.hover}
                line1Property='displayName'
                line2Property='jobTitle'
                line3Property='department'
                view={PersonViewType.threelines}
                className={styles.person}
                disableImageFetch
              />
            </div>
          ))}
        </div>
      ) : (
        <div style={{ padding: '0 4px' }}>
          <Text variant='mediumPlus' as='span'>
            No results found
          </Text>
        </div>
      )}
    </div>
  );
};

export default Results;
