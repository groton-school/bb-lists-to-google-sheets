const App = {
  PREFIX: 'org.groton.BbListsToGoogleSheets',
  LOGO_URL:
    'https://drive.google.com/uc?id\u003d1sany_QufWim04ZXwj1cMWh03E-cYHWk0',

  launch(event) {
    if (event) {
      App.launchEvent = event;
    }
    State.reset();
    State.setFolder(Drive.inferFolder(App.launchEvent));
    // TODO can also infer folder from parent of current spreadsheet
    return App.cards.home();
  },

  actions: {
    home({ parameters: { state } }) {
      State.reset(state);
      return TerseCardService.replaceStack(App.cards.home());
    },
  },

  cards: {
    home() {
      if (State.getSpreadsheet()) {
        return Sheets.cards.options();
      } else {
        return Lists.cards.lists();
      }
    },

    error(message = 'An error occurred') {
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(message))
        .addSection(
          CardService.newCardSection()
            .addWidget(
              TerseCardService.newDecoratedText(
                'State',
                JSON.stringify(
                  JSON.parse(State.toJSON()),
                  null,
                  2
                )
              )
            )
            .addWidget(
              TerseCardService.newTextButton(
                'Start Over',
                '__App_actions_home'
              )
            )
        )
        .build();
    },
  },
};

function __App_launch(e) {
  return App.launch(e);
}

function __App_actions_home(...args) {
  return App.actions.home(...args);
}
