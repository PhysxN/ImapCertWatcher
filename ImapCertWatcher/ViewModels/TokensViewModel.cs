using ImapCertWatcher.Models;
using ImapCertWatcher.Services;
using ImapCertWatcher.UI;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;

namespace ImapCertWatcher.ViewModels
{
    public class TokensViewModel : INotifyPropertyChanged
    {
        private readonly TokenService _service;
        private string _searchText = "";
        private bool _showBusyTokens;

        public bool ShouldShowBusyTokensSection =>
    _showBusyTokens || !string.IsNullOrWhiteSpace(_searchText);

        public ObservableCollection<TokenRecord> FreeTokens { get; } =
            new ObservableCollection<TokenRecord>();

        public ObservableCollection<TokenRecord> BusyTokens { get; } =
            new ObservableCollection<TokenRecord>();

        public ICommand ReloadCommand { get; }
        public ICommand UnassignCommand { get; }
        public ICommand DeleteCommand { get; }

        public string SearchText
        {
            get => _searchText;
            set
            {
                var newValue = value ?? "";

                if (_searchText == newValue)
                    return;

                _searchText = newValue;
                OnPropertyChanged(nameof(SearchText));
                OnPropertyChanged(nameof(ShouldShowBusyTokensSection));
                ApplyFilter();
            }
        }

        public bool ShowBusyTokens
        {
            get => _showBusyTokens;
            set
            {
                if (_showBusyTokens == value)
                    return;

                _showBusyTokens = value;
                OnPropertyChanged(nameof(ShowBusyTokens));
                OnPropertyChanged(nameof(ShouldShowBusyTokensSection));
                ApplyFilter();
            }
        }

        public TokensViewModel(TokenService service)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));

            _service.TokensChanged += ApplyFilter;
            ApplyFilter();

            ReloadCommand = new RelayCommand(async _ => await _service.Reload());

            UnassignCommand = new RelayCommand(async t =>
            {
                if (t is TokenRecord token)
                    await _service.Unassign(token.Id);
            });

            DeleteCommand = new RelayCommand(async t =>
            {
                if (t is TokenRecord token)
                    await _service.Delete(token.Id);
            });
        }

        private bool MatchesSearch(TokenRecord token)
        {
            if (token == null)
                return false;

            if (string.IsNullOrWhiteSpace(_searchText))
                return true;

            var terms = _searchText
                .ToLowerInvariant()
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            var haystack = string.Join(" ",
                token.Sn ?? "",
                token.OwnerFio ?? "",
                token.OwnerFioFormatted ?? "")
                .ToLowerInvariant();

            return terms.All(term => haystack.Contains(term));
        }

        private void ApplyFilter()
        {
            FreeTokens.Clear();
            BusyTokens.Clear();

            bool forceShowBusyBySearch = !string.IsNullOrWhiteSpace(_searchText);
            bool showBusy = _showBusyTokens || forceShowBusyBySearch;

            foreach (var t in _service.Tokens)
            {
                if (t == null)
                    continue;

                if (!MatchesSearch(t))
                    continue;

                if (t.OwnerCertId == null)
                    FreeTokens.Add(t);
                else if (showBusy)
                    BusyTokens.Add(t);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}