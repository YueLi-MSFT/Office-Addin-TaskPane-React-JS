import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import { selectInsertionByHost } from "../../host-relative-text-insertion";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const TextInsertionForOutlook = () => {
  const [text, setText] = useState("");

  const handleTextInsertion = async () => {
    setText(Office.context.mailbox.item.subject);
  };

  const handleTextChange = async (event) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.instructions}>Click the button to read subject.</Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Run
      </Button>
      {text}
    </div>
  );
};

export default TextInsertionForOutlook;
