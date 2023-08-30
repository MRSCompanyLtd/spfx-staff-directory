import * as React from 'react';
import styles from './StaffDirectory.module.scss';
import { IStaffDirectoryProps } from './IStaffDirectoryProps';

import useSearch from '../hooks/useSearch';
import Search from './Search/Search';
import Results from './Results/Results';
import Paging from './Paging/Paging';
import { IPerson } from '../interfaces/IPerson';
import { Dropdown, IDropdownOption, Text } from 'office-ui-fabric-react';
import './global.css';

interface IStaffDirectoryState {
  search: string;
  page: number;
  items: IPerson[];
  selected: string | number;
}

const StaffDirectory: React.FC<IStaffDirectoryProps> = ({
  title,
  isDarkTheme,
  showDepartmentFilter,
  departments,
  group,
  userDisplayName,
  pageSize,
  context
}) => {
  const [state, setState] = React.useState<IStaffDirectoryState>({
    search: '',
    page: 1,
    items: [],
    selected: ''
  });

  const { total, searchByText, getNextPage, loading, results } =
    useSearch(context, group, pageSize);

  const handleChange = React.useCallback(
    (
      event?: React.ChangeEvent<HTMLInputElement> | undefined,
      newValue?: string | undefined
    ) => {
      setState((s) => {
        return {
          ...s,
          search: newValue || ''
        };
      });
    },
    []
  );

  const resetSearch = React.useCallback(() => {
    setState((s) => {
      return {
        ...s,
        search: ''
      };
    });
  }, []);

  const handleSubmit = React.useCallback(async () => {
    const { search, selected } = state;

    setState((s) => ({
      ...s,
      items: [],
      page: 1
    }));

    await searchByText(search, selected.toString()).catch((e) =>
      console.error(e)
    );
  }, [searchByText, state]);

  const goToPage = React.useCallback(
    async (v: number) => {
      if (v > state.page) {
        await getNextPage();
      }

      setState((s) => ({
        ...s,
        page: v
      }));
    },
    [getNextPage, state]
  );

  const handleDropdown = React.useCallback(
    (
      event: React.FormEvent<HTMLDivElement>,
      option?: IDropdownOption | undefined,
      index?: number | undefined
    ) => {
      setState((s) => ({
        ...s,
        selected: option?.key ?? ''
      }));
    },
    []
  );

  const load = React.useCallback(async () => {
    await searchByText('', state.selected as string).catch(e => console.error(e));
  }, [group, state]);

  React.useEffect(() => {
    setState((s) => ({
      ...s,
      items: [],
      search: '',
      page: 1
    }));

    load().catch((e) => console.error(e));
  }, [group, state.selected]);

  React.useEffect(() => {
    setState((s) => ({
      ...s,
      items: [...s.items, ...results]
    }));
  }, [results]);

  const displayedItems = React.useMemo(() => {
    const { page, items } = state;
    const startIndex = (page - 1) * pageSize;
    const endIndex = page * pageSize;

    return items.slice(startIndex, endIndex);
  }, [state, pageSize]);

  return (
    <section className={styles.staffDirectory}>
      <div>
        <div className={styles.title}>
          <Text as='h1' variant='xLarge'>
            {title}
          </Text>
        </div>
        <div className={styles.search}>
          <Search
            placeholder='Search name, department, or job title'
            onChange={handleChange}
            onClear={resetSearch}
            onSearch={handleSubmit}
            value={state.search}
            className={styles.searchBar}
          />
          {showDepartmentFilter && (
            <Dropdown
              options={[
                {
                  key: '',
                  text: 'All departments'
                },
                ...departments
              ]}
              placeholder='Select department'
              selectedKey={state.selected}
              onChange={handleDropdown}
              className={styles.dropdown}
            />
          )}
        </div>
        <br style={{ margin: '2px 0' }} />
        <Paging
          count={total}
          page={state.page}
          pageSize={pageSize}
          onPageChange={goToPage}
        />
        <br style={{ margin: '6px 0' }} />
        <Results results={displayedItems} loading={loading} />
      </div>
    </section>
  );
};

export default StaffDirectory;
