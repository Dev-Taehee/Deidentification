import torch
import torch.nn as nn

from transformers import BertForTokenClassification, logging

logging.set_verbosity_error()


class BERT_NER(nn.Module):
    
    def __init__(self):
        super(BERT_NER, self).__init__()

        self.tag_list = ["[PAD]", "[CLS]", "[SEP]", "O", "PER-B", "PER-I", "ORG-B", "ORG-I",
                         "LOC-B", "LOC-I", "DAT-B", "DAT-I", "TIM-B", "TIM-I", "NUM-B", "NUM-I"]

        self.encoder = BertForTokenClassification.from_pretrained("klue/bert-base", num_labels=len(self.tag_list))
    
    def forward(self, data):
        input_ids, labels, attention_mask = data
        outputs = self.encoder(input_ids, attention_mask=attention_mask, labels=labels)

        return outputs.loss
    
    def predict(self, input_ids):
        # only for single batch
        logits = self.encoder(input_ids).logits
        softmax = torch.softmax(logits, dim=-1)
        pred = torch.argmax(softmax, dim=-1)
        return pred
