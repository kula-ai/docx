# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'

describe Docx::Elements::Text do
  # Builds a bare <w:t> node so #content= whitespace handling can be exercised
  # without a full .docx fixture.
  def text_from(inner)
    xml = <<~XML
      <w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">#{inner}</w:t>
    XML
    node = Nokogiri::XML(xml).root
    [described_class.new(node), node]
  end

  describe '#content=' do
    # Word/OOXML trims a <w:t>'s leading/trailing whitespace unless it carries
    # xml:space="preserve". Consolidation can leave a run holding edge whitespace
    # (e.g. " office." after a merged token), so #content= must re-assert it.
    it 'sets xml:space="preserve" when the new content has a leading space' do
      text, node = text_from('offer.office')
      text.content = ' office.'

      expect(node.attribute('space')&.value).to eq('preserve')
    end

    it 'sets xml:space="preserve" when the new content has a trailing space' do
      text, node = text_from('offer.name')
      text.content = 'Grade '

      expect(node.attribute('space')&.value).to eq('preserve')
    end

    it 'does not add xml:space when the content has no edge whitespace' do
      text, node = text_from('offer.name')
      text.content = 'office.'

      expect(node.attribute('space')).to be_nil
    end

    it 'does not add xml:space for empty content' do
      text, node = text_from('x')
      text.content = ''

      expect(node.attribute('space')).to be_nil
    end

    it 'writes the content regardless of whitespace' do
      text, node = text_from('placeholder')
      text.content = ' office.'

      expect(node.content).to eq(' office.')
    end
  end
end
