const TerseCardService = {
  replaceStack(card, url = null) {
    var action = CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(card));

    if (url) {
      action = action.setOpenLink(CardService.newOpenLink()
        .setUrl(url));
    }

    return action.build();
  },

  pushCard(card) {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(card))
      .build();
  },

  newCardHeader(title) {
    return CardService.newCardHeader()
      .setTitle(title);
  },

  newTextParagraph(text) {
    return CardService.newTextParagraph()
      .setText(text);
  },

  newDecoratedText(topLabel = null, text = null, bottomLabel = null, wrap = true) {
    return CardService.newDecoratedText()
      .setTopLabel(topLabel || ' ')
      .setText(text || ' ')
      .setWrapText(wrap)
      .setBottomLabel(bottomLabel || ' ')
  },

  newTextButton(text, functionName, stateChamge = null) {
    return CardService.newTextButton()
      .setText(text)
      .setOnClickAction(TerseCardService.newAction(functionName, stateChamge));
  },

  newAction(functionName, stateChange = null) {
    return CardService.newAction()
      .setFunctionName(functionName)
      .setParameters({ state: State.toJSON(stateChange) });
  }
}
