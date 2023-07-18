import * as React from 'react';
import {
  IDropdownOption,
  Dropdown,
  Spinner
} from 'office-ui-fabric-react';
import { IGroupSelectProps } from './IGroupSelectProps';

const GroupSelect: React.FC<IGroupSelectProps> = ({ ...props }) => {
  const [selected, setSelected] = React.useState<string | number>(
    props.selected
  );
  const [options, setOptions] = React.useState<IDropdownOption[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);

  const loadOptions = React.useCallback(() => {
    setLoading(true);

    props
      .loadOptions()
      .then((opts: IDropdownOption[]) => {
        setOptions(opts);
      })
      .catch((e) => {
        console.error(e);

        setOptions([]);
      });

    setLoading(false);
  }, [props.loadOptions]);

  const handleChange = React.useCallback((
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption | undefined,
    index?: number | undefined
  ) => {
    if (option) {
      setSelected(option.key);
      if (props.onChange) {
        props.onChange(option, index);
      }
    }
  }, [props.onChange]);

  React.useEffect(() => {
    loadOptions();
  }, [loadOptions]);

  return (
    <div>
      <Dropdown
        label={props.label}
        options={options}
        selectedKey={selected}
        disabled={props.disabled}
        onChange={handleChange}
      />
      {loading && <Spinner />}
    </div>
  );
};

export default GroupSelect;
