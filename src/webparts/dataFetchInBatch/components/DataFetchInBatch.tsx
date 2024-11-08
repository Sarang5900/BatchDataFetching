import * as React from "react";
import { sp } from "@pnp/sp/presets/all";
import {
  DetailsList,
  DetailsListLayoutMode,
  DetailsRow,
  IColumn,
  IDetailsRowProps,
  IDetailsRowStyles,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Dropdown,
  IDropdownOption,
  SearchBox,
} from "@fluentui/react";
import { useState, useEffect } from "react";
import { IDataFetchInBatchProps } from "./IDataFetchInBatchProps";

interface IListItem {
  Title: string;
  field_1: string; // name
  field_2: string; // surname
  field_3: string; // gender
  field_4: string; // country
  field_5: number; // age
  field_6: string; // date 
  field_7: string; // id
}

const DataFetchInBatch: React.FC<IDataFetchInBatchProps> = (props) => {
  useEffect(() => {
    sp.setup({
      spfxContext: props.context as any,
    });
  }, [props.context]);

  const [items, setItems] = useState<IListItem[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [hasMoreRecordsNext, setHasMoreRecordsNext] = useState<boolean>(true);
  const [hasMoreRecordsPrev, setHasMoreRecordsPrev] = useState<boolean>(false);
  const [batchSize] = useState<number>(20);
  const [skip, setSkip] = useState<number>(0);
  const [filterValue, setFilterValue] = useState<string | undefined>(undefined);
  const [searchText, setSearchText] = useState<string>("");

  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    const day = ("0" + date.getDate()).slice(-2);
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
  };

  
  const columns: IColumn[] = [
    {
      key: "column1",
      name: "Sr No",
      fieldName: "Title",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column2",
      name: "First Name",
      fieldName: "field_1",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column3",
      name: "Last Name",
      fieldName: "field_2",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column4",
      name: "Gender",
      fieldName: "field_3",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column5",
      name: "Country",
      fieldName: "field_4",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column6",
      name: "Age",
      fieldName: "field_5",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column7",
      name: "Date",
      fieldName: "field_6",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IListItem) => {
        return formatDate(item.field_6);
      },
    },
    {
      key: "column8",
      name: "ID",
      fieldName: "field_7",
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
  ];

  const options: IDropdownOption[] = [
    { key: "all", text: "All" },
    { key: "United States", text: "United States" },
    { key: "Great Britain", text: "Great Britain" },
    { key: "France", text: "France" },
  ];

  const fetchBatch = async (newSkip: number, filter?: string, search?: string): Promise<void> => {
    setLoading(true);
    try {
      const list = sp.web.lists.getByTitle("BatchExcelList");
      let query = list.items
        .select(
          "Title",
          "field_1",
          "field_2",
          "field_3",
          "field_4",
          "field_5",
          "field_6",
          "field_7"
        )
        .top(batchSize)
        .skip(newSkip);

      if (filter) {
        query = query.filter(`field_4 eq '${filter}'`);
      }

      if (search) {
        query = query.filter(`Title eq '${search}' or field_1 eq '${search}' or field_2 eq '${search}' or field_3 eq '${search}' or field_5 eq '${search}' or field_7 eq '${search}'`);
      }
      

      const response = await query.get();

      console.log("Fetched Data in Batch:", response);

      if (response.length > 0) {
        setItems(response);
        setSkip(newSkip);
        setHasMoreRecordsPrev(newSkip > 0);
        setHasMoreRecordsNext(response.length === batchSize);
      } else {
        setHasMoreRecordsNext(false);
      }
    } catch (error) {
      console.error("Error in batch fetching:", error);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchBatch(0, filterValue, searchText).catch((error) =>
      console.error("Error during initial data load:", error)
    );
  }, [filterValue, searchText]);

  const handleFilterChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    setFilterValue(option?.key === "all" ? undefined : (option?.key as string));
    setSkip(0); 
  };

  const handleSearchChange = (newValue?: string) => {
    setSearchText(newValue || "");
    setSkip(0); 
  };

  const loadNextItems = (): void => {
    if (hasMoreRecordsNext) {
      fetchBatch(skip + batchSize, filterValue, searchText).catch((error) =>
        console.error("Error loading next items:", error)
      );
    }
  };

  const loadPreviousItems = (): void => {
    if (hasMoreRecordsPrev && skip - batchSize >= 0) {
      fetchBatch(skip - batchSize, filterValue, searchText).catch((error) =>
        console.error("Error loading previous items:", error)
      );
    }
  };

  const onRenderRow = (props: IDetailsRowProps): JSX.Element | null => {
    if (!props) return null;

    const rowStyles: IDetailsRowStyles = {
      root: {
        backgroundColor: props.itemIndex % 2 === 0 ? "#f3f4f6" : "#ffffff",
      },
      cell: undefined,
      cellAnimation: undefined,
      cellUnpadded: undefined,
      cellPadded: undefined,
      checkCell: undefined,
      isRowHeader: undefined,
      isMultiline: undefined,
      fields: undefined,
      cellMeasurer: undefined,
      check: undefined
    };

    return <DetailsRow {...props} styles={rowStyles} />;
  };

  return (
    <div>
      <SearchBox
        placeholder="Search..."
        onChange={(_, newValue) => handleSearchChange(newValue)}
        styles={{ root: { marginBottom: 20, width: 200 } }}
      />
      
      <Dropdown
        label="Filter by Country"
        options={options}
        onChange={handleFilterChange}
        selectedKey={filterValue || "all"}
        styles={{ root: { width: 200, marginBottom: "20px" } }}
      />

      <DetailsList
        items={items}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.fixedColumns}
        onRenderRow={onRenderRow}
      />
      <div style={{ display: "flex", justifyContent: "center", marginTop: "20px" }}>
        {loading ? (
          <Spinner size={SpinnerSize.medium} label="Loading..." />
        ) : (
          <>
            <PrimaryButton
              onClick={loadPreviousItems}
              text="Previous"
              disabled={!hasMoreRecordsPrev}
              style={{ marginRight: "10px" }}
            />
            <PrimaryButton
              onClick={loadNextItems}
              text="Next"
              disabled={!hasMoreRecordsNext}
            />
          </>
        )}
      </div>
      {!hasMoreRecordsNext && !loading && (
        <div style={{ textAlign: "center", color: "gray", marginTop: "10px" }}>
          Cannot load more records
        </div>
      )}
    </div>
  );
};

export default DataFetchInBatch;
