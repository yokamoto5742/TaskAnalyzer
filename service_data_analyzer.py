import polars as pl


class TaskDataAnalyzer:
    """
    タスクデータの分析と集計を行うクラス
    """
    @staticmethod
    def create_dataframes(tasks, daily_tasks, communication_tasks, all_items):
        """
        リストからpolarsデータフレームを作成する
        
        Parameters:
        -----------
        tasks : list
            タスクのリスト
        daily_tasks : list
            デイリータスクのリスト
        communication_tasks : list
            コミュニケーションタスクのリスト
        all_items : list
            全項目のリスト
            
        Returns:
        --------
        tuple
            (df, daily_df, comm_df, all_items_df)
        """
        df = pl.DataFrame(tasks)
        daily_df = pl.DataFrame(daily_tasks)
        comm_df = pl.DataFrame(communication_tasks)
        all_items_df = pl.DataFrame(all_items)
        
        return df, daily_df, comm_df, all_items_df

    @staticmethod
    def aggregate_dataframe(data_frame, group_by_col='content', filter_condition=None):
        """
        データフレームを集計する
        
        Parameters:
        -----------
        data_frame : DataFrame
            集計対象のデータフレーム
        group_by_col : str
            グループ化するカラム名
        filter_condition : Expression
            フィルター条件
            
        Returns:
        --------
        DataFrame
            集計結果のデータフレーム
        """
        if filter_condition is not None:
            data_frame = data_frame.filter(filter_condition)

        if group_by_col == 'content' and 'name' in data_frame.columns:
            return (
                data_frame.group_by(['name', group_by_col])
                .agg([
                    pl.col('minutes').sum().alias('total_minutes'),
                    (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
                    pl.col('minutes').count().alias('frequency')
                ])
                .sort(['name', 'total_minutes'], descending=[False, True])
            )

        return (
            data_frame.group_by(group_by_col)
            .agg([
                pl.col('minutes').sum().alias('total_minutes'),
                (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
                pl.col('minutes').count().alias('frequency')
            ])
            .sort('total_minutes', descending=True)
        )

    def analyze_task_data(self, tasks, daily_tasks, comm_tasks, all_items):
        """
        タスクデータを分析する
        
        Parameters:
        -----------
        tasks : list
            タスクのリスト
        daily_tasks : list
            デイリータスクのリスト
        comm_tasks : list
            コミュニケーションタスクのリスト
        all_items : list
            全項目のリスト
            
        Returns:
        --------
        tuple
            (clerk_tasks, non_clerk_tasks, daily_tasks_agg, 
             communication_by_name, communication_by_content, all_items_summary)
        """
        df, daily_df, comm_df, all_items_df = self.create_dataframes(
            tasks, daily_tasks, comm_tasks, all_items
        )
        
        # クラーク業務の集計
        clerk_tasks = self.aggregate_dataframe(
            df,
            filter_condition=pl.col('content').str.contains('クラーク業務')
        )
        
        # クラーク業務以外の集計
        non_clerk_tasks = self.aggregate_dataframe(
            df,
            filter_condition=~pl.col('content').str.contains('クラーク業務')
        )
        
        # デイリータスクの集計
        daily_tasks_agg = self.aggregate_dataframe(daily_df)
        
        # コミュニケーション（名前別）の集計
        communication_by_name = self.aggregate_dataframe(comm_df, group_by_col='name')
        
        # コミュニケーション（内容別）の集計
        communication_by_content = (
            comm_df.group_by(['content', 'name'])
            .agg([
                pl.col('minutes').sum().alias('total_minutes'),
                (pl.col('minutes').sum() / 60).cast(pl.Int64).alias('total_hours'),
                pl.col('minutes').count().alias('frequency')
            ])
            .sort(['name', 'total_minutes'], descending=[False, True])
            .select(['content', 'name', 'total_minutes', 'total_hours', 'frequency'])
        )
        
        # 全項目の集計
        all_items_summary = self.aggregate_dataframe(all_items_df)
        
        return (
            clerk_tasks,
            non_clerk_tasks,
            daily_tasks_agg,
            communication_by_name,
            communication_by_content,
            all_items_summary
        )
