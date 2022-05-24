from ipywidgets import interact, Dropdown


class IActiveDF(object):

    def __init__(self, df_in, col=1, by_col=None, des='E-mails'):
        self.df_in = df_in
        self.by_col = by_col
        self.col = col
        self.des = des
        self.value = ''
        self.get_dropdown()

    def get_dropdown(self):
        df_cols = self.df_in.columns.values.tolist()
        if self.by_col is None:
            options = self.df_in[df_cols[self.col]].tolist()
        else:
            options =\
                self.df_in.sort_values(by=df_cols[self.by_col],
                                       ascending=False)[df_cols[self.col]].tolist()
        df_ls = {self.des: options}
        dd_ = Dropdown(options=df_ls[self.des], description=self.des)

        @interact(value=dd_)
        def print_value(value):
            self.value = value
