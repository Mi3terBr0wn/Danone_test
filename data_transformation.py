import pandas as pd
from datetime import datetime
from openpyxl.styles import is_date_format


def replace_value_in_column_by_condition_on_equality(df, df_column, old_condition, new_condition):
    df.loc[events_df[df_column] == old_condition, df_column] = new_condition


def customer_churn_rate(customer_count_beginning_of_month_f, customer_count_end_of_month_f):
    if customer_count_beginning_of_month_f != 0:
        return ((customer_count_beginning_of_month_f - customer_count_end_of_month_f) /
                customer_count_beginning_of_month_f).astype(float)
    else:
        return pd.NA


def create_df_with_grouping(df, df_column):
    unique_column_values = df[df_column].unique()
    new_df = pd.DataFrame(columns=df.columns[1:-2])
    new_df.insert(loc=len(new_df.columns), column=df_column, value=unique_column_values)
    return new_df


def customer_churn_rate_df_with_grouping(by_group_customer_churn_rate_df, group_column):
    for col in events_df.columns:
        if col not in ['Customer ID', 'Country', 'Billing Plan']:
            customer_count_beginning_of_month_with_grouping = (
                events_df[events_df[col].isin([2, 3, 4])].groupby(group_column)[col].count())
            customer_count_end_of_month_with_grouping = events_df[events_df[col] == 2].groupby(group_column)[
                col].count()

            group_customer_churn_rate = ((customer_count_beginning_of_month_with_grouping -
                                          customer_count_end_of_month_with_grouping) /
                                         customer_count_beginning_of_month_with_grouping.replace(0, pd.NA))

            by_group_customer_churn_rate_df[col] = by_group_customer_churn_rate_df[
                group_column].map(group_customer_churn_rate).fillna(0)
    return by_group_customer_churn_rate_df


if __name__ == "__main__":
    excel_file_path = '03_Churn_Rate_Calculation_Задание_Advanced_Excel.xlsx'
    excel_reader = pd.ExcelFile(excel_file_path)
    events_df = excel_reader.parse('02_Subscription_Events')
    events_df = events_df.reset_index()

    replace_value_in_column_by_condition_on_equality(events_df, 'Event Type', 'Subscribed', 1)
    replace_value_in_column_by_condition_on_equality(events_df, 'Event Type', 'Unsubscribed', 3)

    events_df['Event Type'] = events_df['Event Type'].astype(int)

    events_customers_df = events_df[['index', 'Customer ID']].copy()

    events_df['Month'] = events_df['Date'].apply(lambda x: x.replace(day=1)).dt.date
    events_df.drop('Date', axis=1, inplace=True)
    events_df.drop_duplicates(['Customer ID', 'Event Type', 'Month'])
    events_df = events_df.sort_values('Month')

    events_df = events_df.pivot(index='index', columns='Month', values='Event Type')
    events_df = events_df.fillna(0)
    events_df.columns = [str(col) for col in events_df.columns]
    events_df.columns = ['.'.join(col.split('-')[::-1][1:]) for col in events_df.columns]

    events_df = events_df.merge(events_customers_df, on=['index', 'index'], how='left')
    events_df.drop('index', axis=1, inplace=True)

    events_df = events_df.groupby('Customer ID').sum().reset_index()
    events_df.set_index('Customer ID')

    for i in range(1, len(events_df.columns)):
        current_col = events_df.columns[i]
        prev_col = events_df.columns[i - 1]
        mask = (events_df[prev_col].isin([1, 2])) & (events_df[current_col] != 3)
        events_df.loc[mask, current_col] = 2

    for i in range(len(events_df.columns) - 2, 0, -1):
        current_col = events_df.columns[i]
        prev_col = events_df.columns[i + 1]
        mask = (events_df[prev_col].isin([2, 3])) & (events_df[current_col] != 1)
        events_df.loc[mask, current_col] = 2

    overall_customer_churn_rate_df = pd.DataFrame(columns=events_df.columns[1:])

    for column in events_df.columns[1:]:
        customer_count_beginning_of_month = events_df[column].isin([2, 3, 4]).count()
        customer_count_end_of_month = events_df[column].eq(2).sum()
        overall_customer_churn_rate_df.at[0, column] = (
            customer_churn_rate(customer_count_beginning_of_month.astype(float),
                                customer_count_end_of_month.astype(float)))


    customers_df = excel_reader.parse('01_Customers', index_col=0)

    events_df = events_df.merge(customers_df, on=['Customer ID', 'Customer ID'], how='left')

    by_country_customer_churn_rate_df = pd.DataFrame({'Country': events_df['Country'].unique()})
    by_country_customer_churn_rate_df = (
        customer_churn_rate_df_with_grouping(by_country_customer_churn_rate_df, 'Country'))

    by_billing_plan_customer_churn_rate_df = pd.DataFrame({'Billing Plan': events_df['Billing Plan'].unique()})
    by_billing_plan_customer_churn_rate_df = (
        customer_churn_rate_df_with_grouping(by_billing_plan_customer_churn_rate_df, 'Billing Plan'))

    with pd.ExcelWriter('Сustomer_churn_rate_file.xlsx', engine='xlsxwriter') as writer:
        overall_customer_churn_rate_df.to_excel(writer, sheet_name='Без группировок')
        by_country_customer_churn_rate_df.to_excel(writer, sheet_name='По странам')
        by_billing_plan_customer_churn_rate_df.to_excel(writer, sheet_name='По типу подписки')
