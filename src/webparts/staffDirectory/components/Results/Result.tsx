import { Person, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import { PersonCardInteraction, PersonViewType } from '@microsoft/mgt-spfx';
import * as React from 'react';
import styles from './Results.module.scss';

interface IResultProps extends MgtTemplateProps {}

const Result: React.FC<IResultProps> = ({ dataContext }) => {
  const { person } = dataContext;

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
    <div className={styles.result}>
      <Person
        userId={person.id}
        personDetails={{
          ...person,
          personImage: getPhoto(person.picture ?? '')
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
  );
};

export default Result;
