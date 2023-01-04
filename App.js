const App = {
  LOGO_URL: 'https://drive.google.com/uc?id\u003d1sany_QufWim04ZXwj1cMWh03E-cYHWk0',
  launchEvent: null,

  launch(event) {
    if (event) {
      App.launchEvent = event;
    }
    State.reset()
    State.folder = Drive.inferFolder(App.launchEvent);
    if (State.sheet) {
      return Sheets.cards.options();
    } else {
      return Lists.cards.lists();
    }
  },

  actions: {
    home() {
      var url = null;
      var card = null;
      if (arguments && arguments.length > 0) {
        url = arguments[0].url;
        card = arguments[0].card;
      }

      var action = CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .popToRoot()
          .updateCard(card || launch()));

      /*
       * FIXME Opening link not working
       *   Related to #1 probably
       *   Order of operations doesn't seem to matter (OpenLink before Navigation or vice versa)
       */
      if (url) {
        action = action.setOpenLink(CardService.newOpenLink()
          .setUrl(url));
      }
      return action.build();
    },

    error({ parameters: { state, message } }) {
      State.restore(state);
      return App.actions.home({ card: App.cards.error(message) });
    }
  },

  cards: {
    error(message = "An error occurred") {
      return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(message))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newDecoratedText()
            .setTopLabel('State')
            .setText(JSON.stringify(JSON.parse(State.toJSON()), null, 2))
            .setWrapText(true))
          .addWidget(CardService.newTextButton()
            .setText('Start Over')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__App_actions_home'))))
        .build();
    }
  }
}

function launch(e) {
  return App.launch(e);
}

function __App_actions_home(...args) {
  return App.actions.home(...args);
}
