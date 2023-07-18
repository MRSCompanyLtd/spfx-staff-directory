import * as React from 'react';
import { IPagingProps } from './IPagingProps';
import styles from './Paging.module.scss';
import { IconButton } from 'office-ui-fabric-react';

const Paging: React.FC<IPagingProps> = ({
  count,
  page,
  pageSize,
  onPageChange
}) => {
  const selected: React.MutableRefObject<number> = React.useRef(page);

  const pageLength = count === 0 ? 1 : Math.ceil(count / pageSize);

  const handleChange = React.useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      const num: number = Number(e.currentTarget.ariaValueNow);
      if (!Number.isNaN(num)) {
        selected.current = num;
        onPageChange(num);
      }
    },
    [onPageChange]
  );

  const clickNext = React.useCallback(() => {
    selected.current = Math.min(selected.current + 1, pageLength);
    onPageChange(selected.current);
  }, [pageLength, onPageChange]);

  const clickBack = React.useCallback(() => {
    selected.current = Math.max(selected.current - 1, 1);
    onPageChange(selected.current);
  }, [onPageChange]);

  React.useEffect(() => {
    selected.current = 1;
  }, [count, pageSize]);

  const pages = (): JSX.Element => {
    const elements: JSX.Element[] = [];
    const maxPageItems = 5; // Set the maximum number of page items to display (both preceding and succeeding)

    let startPage = Math.max(1, page - Math.floor(maxPageItems / 2));
    const endPage = Math.min(pageLength, startPage + maxPageItems - 1);

    if (endPage - startPage < maxPageItems - 1) {
      // Adjust startPage if the range has fewer items than the desired maxPageItems
      startPage = Math.max(1, endPage - maxPageItems + 1);
    }

    for (let i = startPage; i <= endPage; i++) {
      elements.push(
        <IconButton
          key={i}
          aria-valuenow={i}
          className={`${styles.page} ${selected.current === i && styles.active}`}
          onClick={handleChange}
        >
          {i}
        </IconButton>
      );
    }

    return <div className={styles.pages}>{elements}</div>;
  };

  return (
    <div className={styles.paging}>
      <div className={styles.count}>
        {count > 0 ? `There are ${count} results` : `No results`}
      </div>
      <div className={styles.pageActions} style={{ display: count === 0 ? 'none' : 'flex' }}>
        {pages()}
        <IconButton
            onClick={clickBack}
            disabled={selected.current === 1}
            iconProps={{ iconName: 'Back' }}
        />
        <IconButton
          onClick={clickNext}
          disabled={selected.current === pageLength}
          iconProps={{ iconName: 'Forward' }}
        >
          Next
        </IconButton>
      </div>
    </div>
  );
};

export default Paging;
