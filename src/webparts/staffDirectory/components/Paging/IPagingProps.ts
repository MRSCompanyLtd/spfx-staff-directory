export interface IPagingProps {
    count: number;
    page: number;
    pageSize: number;
    onPageChange: (val: number) => void;
}