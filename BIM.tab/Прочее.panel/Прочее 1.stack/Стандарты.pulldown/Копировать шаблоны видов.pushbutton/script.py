# coding=utf-8
"""Copy selected view templates to other open models."""
# pylint: disable=import-error,invalid-name
from pyrevit import revit
from pyrevit import forms
from pyrevit.forms import *

from dosymep_libs.bim4everyone import *
from pyrevit import HOST_APP, EXEC_PARAMS, DOCS, BIN_DIR

DEFAULT_INPUTWINDOW_WIDTH = 500


class SelectFromList(TemplateUserInputWindow):
    xaml_source = 'SelectFromList.xaml'

    @property
    def use_regex(self):
        """Is using regex?"""
        return self.regexToggle_b.IsChecked

    def _setup(self, **kwargs):
        # custom button name?
        button_name = kwargs.get('button_name', 'Select')
        if button_name:
            self.select_b.Content = button_name

        # attribute to use as name?
        self._nameattr = kwargs.get('name_attr', None)

        # multiselect?
        if kwargs.get('multiselect', False):
            self.multiselect = True
            self.list_lb.SelectionMode = Controls.SelectionMode.Extended
            self.show_element(self.checkboxbuttons_g)
        else:
            self.multiselect = False
            self.list_lb.SelectionMode = Controls.SelectionMode.Single
            self.hide_element(self.checkboxbuttons_g)

        # info panel?
        self.info_panel = kwargs.get('info_panel', False)

        # return checked items only?
        self.return_all = kwargs.get('return_all', False)

        # filter function?
        self.filter_func = kwargs.get('filterfunc', None)

        # reset function?
        self.reset_func = kwargs.get('resetfunc', None)
        if self.reset_func:
            self.show_element(self.reset_b)

        # context group title?
        self.ctx_groups_title = \
            kwargs.get('group_selector_title', 'List Group')
        self.ctx_groups_title_tb.Text = self.ctx_groups_title

        self.ctx_groups_active = kwargs.get('default_group', None)

        # check for custom templates
        items_panel_template = kwargs.get('items_panel_template', None)
        if items_panel_template:
            self.Resources["ItemsPanelTemplate"] = items_panel_template

        item_container_template = kwargs.get('item_container_template', None)
        if item_container_template:
            self.Resources["ItemContainerTemplate"] = item_container_template

        item_template = kwargs.get('item_template', None)
        if item_template:
            self.Resources["ItemTemplate"] = \
                item_template

        # nicely wrap and prepare context for presentation, then present
        self._prepare_context()

        # setup search and filter fields
        self.hide_element(self.clrsearch_b)

        # active event listeners
        self.search_tb.TextChanged += self.search_txt_changed
        self.ctx_groups_selector_cb.SelectionChanged += self.selection_changed

        self.clear_search(None, None)

    def _prepare_context_items(self, ctx_items):
        new_ctx = []
        # filter context if necessary
        if self.filter_func:
            ctx_items = filter(self.filter_func, ctx_items)

        for item in ctx_items:
            if isinstance(item, TemplateListItem):
                item.checkable = self.multiselect
                new_ctx.append(item)
            else:
                new_ctx.append(
                    TemplateListItem(item,
                                     checkable=self.multiselect,
                                     name_attr=self._nameattr)
                )

        return new_ctx

    def _prepare_context(self):
        if isinstance(self._context, dict) and self._context.keys():
            self._update_ctx_groups(sorted(self._context.keys()))
            new_ctx = {}
            for ctx_grp, ctx_items in self._context.items():
                new_ctx[ctx_grp] = self._prepare_context_items(ctx_items)
            self._context = new_ctx
        else:
            self._context = self._prepare_context_items(self._context)

    def _update_ctx_groups(self, ctx_group_names):
        self.show_element(self.ctx_groups_dock)
        self.ctx_groups_selector_cb.ItemsSource = ctx_group_names
        if self.ctx_groups_active in ctx_group_names:
            self.ctx_groups_selector_cb.SelectedIndex = \
                ctx_group_names.index(self.ctx_groups_active)
        else:
            self.ctx_groups_selector_cb.SelectedIndex = 0

    def _get_active_ctx_group(self):
        return self.ctx_groups_selector_cb.SelectedItem

    def _get_active_ctx(self):
        if isinstance(self._context, dict):
            return self._context[self._get_active_ctx_group()]
        else:
            return self._context

    def _list_options(self, option_filter=None):
        if option_filter:
            self.checkall_b.Content = 'Выбрать все'
            self.uncheckall_b.Content = 'Сбросить выделение'
            self.toggleall_b.Content = 'Инвертировать'
            # get a match score for every item and sort high to low
            fuzzy_matches = sorted(
                [(x,
                  coreutils.fuzzy_search_ratio(
                      target_string=x.name,
                      sfilter=option_filter,
                      regex=self.use_regex))
                 for x in self._get_active_ctx()],
                key=lambda x: x[1],
                reverse=True
            )
            # filter out any match with score less than 80
            self.list_lb.ItemsSource = \
                ObservableCollection[TemplateListItem](
                    [x[0] for x in fuzzy_matches if x[1] >= 80]
                )
        else:
            self.checkall_b.Content = 'Выбрать все'
            self.uncheckall_b.Content = 'Сбросить выделение'
            self.toggleall_b.Content = 'Инвертировать'
            self.list_lb.ItemsSource = \
                ObservableCollection[TemplateListItem](self._get_active_ctx())

    @staticmethod
    def _unwrap_options(options):
        unwrapped = []
        for optn in options:
            if isinstance(optn, TemplateListItem):
                unwrapped.append(optn.unwrap())
            else:
                unwrapped.append(optn)
        return unwrapped

    def _get_options(self):
        if self.multiselect:
            if self.return_all:
                return [x for x in self._get_active_ctx()]
            else:
                return self._unwrap_options(
                    [x for x in self._get_active_ctx()
                     if x.state or x in self.list_lb.SelectedItems]
                )
        else:
            return self._unwrap_options([self.list_lb.SelectedItem])[0]

    def _set_states(self, state=True, flip=False, selected=False):
        if selected:
            current_list = self.list_lb.SelectedItems
        else:
            current_list = self.list_lb.ItemsSource
        for checkbox in current_list:
            # using .checked to push ui update
            if flip:
                checkbox.checked = not checkbox.checked
            else:
                checkbox.checked = state

    def _toggle_info_panel(self, state=True):
        if state:
            # enable the info panel
            self.splitterCol.Width = System.Windows.GridLength(8)
            self.infoCol.Width = System.Windows.GridLength(self.Width / 2)
            self.show_element(self.infoSplitter)
            self.show_element(self.infoPanel)
        else:
            self.splitterCol.Width = self.infoCol.Width = \
                System.Windows.GridLength.Auto
            self.hide_element(self.infoSplitter)
            self.hide_element(self.infoPanel)

    def toggle_all(self, sender, args):  # pylint: disable=W0613
        """Handle toggle all button to toggle state of all check boxes."""
        self._set_states(flip=True)

    def check_all(self, sender, args):  # pylint: disable=W0613
        """Handle check all button to mark all check boxes as checked."""
        self._set_states(state=True)

    def uncheck_all(self, sender, args):  # pylint: disable=W0613
        """Handle uncheck all button to mark all check boxes as un-checked."""
        self._set_states(state=False)

    def check_selected(self, sender, args):  # pylint: disable=W0613
        """Mark selected checkboxes as checked."""
        self._set_states(state=True, selected=True)

    def uncheck_selected(self, sender, args):  # pylint: disable=W0613
        """Mark selected checkboxes as unchecked."""
        self._set_states(state=False, selected=True)

    def button_reset(self, sender, args):  # pylint: disable=W0613
        if self.reset_func:
            all_items = self.list_lb.ItemsSource
            self.reset_func(all_items)

    def button_select(self, sender, args):  # pylint: disable=W0613
        """Handle select button click."""
        self.response = self._get_options()
        self.Close()

    def search_txt_changed(self, sender, args):  # pylint: disable=W0613
        """Handle text change in search box."""
        if self.info_panel:
            self._toggle_info_panel(state=False)

        if self.search_tb.Text == '':
            self.hide_element(self.clrsearch_b)
        else:
            self.show_element(self.clrsearch_b)

        self._list_options(option_filter=self.search_tb.Text)

    def selection_changed(self, sender, args):
        if self.info_panel:
            self._toggle_info_panel(state=False)

        self._list_options(option_filter=self.search_tb.Text)

    def selected_item_changed(self, sender, args):
        if self.info_panel and self.list_lb.SelectedItem is not None:
            self._toggle_info_panel(state=True)
            self.infoData.Text = \
                getattr(self.list_lb.SelectedItem, 'description', '')

    def toggle_regex(self, sender, args):
        """Activate regex in search"""
        self.regexToggle_b.Content = \
            self.Resources['regexIcon'] if self.use_regex \
                else self.Resources['filterIcon']
        self.search_txt_changed(sender, args)
        self.search_tb.Focus()

    def clear_search(self, sender, args):  # pylint: disable=W0613
        """Clear search box."""
        self.search_tb.Text = ' '
        self.search_tb.Clear()
        self.search_tb.Focus()


def select_viewtemplates(title='Выберите шаблоны видов',
                         button_name='Выбрать',
                         width=DEFAULT_INPUTWINDOW_WIDTH,
                         multiple=True,
                         filterfunc=None,
                         doc=None):
    doc = doc or DOCS.doc
    all_viewtemplates = revit.query.get_all_view_templates(doc=doc)

    if filterfunc:
        all_viewtemplates = filter(filterfunc, all_viewtemplates)

    selected_viewtemplates = SelectFromList.show(
        sorted([ViewOption(x) for x in all_viewtemplates],
               key=lambda x: x.name),
        title=title,
        button_name=button_name,
        width=width,
        multiselect=multiple,
        checked_only=True
    )

    return selected_viewtemplates


def select_open_docs(title='Select Open Documents',
                     button_name='OK',
                     width=DEFAULT_INPUTWINDOW_WIDTH,  # pylint: disable=W0613
                     multiple=True,
                     check_more_than_one=True,
                     filterfunc=None):
    # find open documents other than the active doc
    open_docs = [d for d in revit.docs if not d.IsLinked]  # pylint: disable=E1101
    if check_more_than_one:
        open_docs.remove(revit.doc)  # pylint: disable=E1101

    if not open_docs:
        alert('Найден только один открытый проект. '
              'Должно быть открыто минимум два проекта. '
              'Операция не выполнена.')
        return

    return SelectFromList.show(
        open_docs,
        name_attr='Title',
        multiselect=multiple,
        title=title,
        button_name=button_name,
        filterfunc=filterfunc
    )


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selected_viewtemplates = select_viewtemplates(doc=revit.doc)

    if selected_viewtemplates:
        dest_docs = select_open_docs(title='Выберите документ назначения')
        if dest_docs:
            for ddoc in dest_docs:
                with revit.Transaction('Copy View Templates', doc=ddoc):
                    revit.create.copy_viewtemplates(
                        selected_viewtemplates,
                        src_doc=revit.doc,
                        dest_doc=ddoc
                    )


script_execute()
