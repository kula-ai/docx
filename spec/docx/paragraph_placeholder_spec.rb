# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'

describe 'Paragraph placeholder consolidation' do
  let(:described_class) { Docx::Elements::Containers::Paragraph }

  # Builds a bare <w:p> node so placeholder consolidation can be exercised
  # without a full .docx fixture.
  def paragraph_from(inner_runs)
    xml = <<~XML
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        #{inner_runs}
      </w:p>
    XML
    node = Nokogiri::XML(xml).root
    described_class.new(node)
  end

  describe '#validate_placeholder_content' do
    # Reproduces the real customer template: Word splits each {{token}} across
    # runs (proofing marks) and fuses one token's closing "}}" into the same run
    # as the next token's opening "{{" (e.g. a run reading "}} at CTC {{"). Two
    # placeholders then share that boundary run.
    let(:shared_boundary_runs) do
      <<~XML
        <w:r><w:t xml:space="preserve">Your title is {{</w:t></w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r><w:t>offer.job_title</w:t></w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r><w:t xml:space="preserve">}} at CTC {{</w:t></w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r><w:t>offer.salary_amount</w:t></w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r><w:t>}}.</w:t></w:r>
      XML
    end

    it 'consolidates without altering the visible text' do
      paragraph = paragraph_from(shared_boundary_runs)

      expect(paragraph.text).to eq('Your title is {{offer.job_title}} at CTC {{offer.salary_amount}}.')
    end

    it 'keeps each placeholder within a single run so substitution leaves no orphan brace' do
      paragraph = paragraph_from(shared_boundary_runs)

      paragraph.each_text_run do |run|
        run.substitute('{{offer.job_title}}', 'Senior Account Executive')
        run.substitute('{{offer.salary_amount}}', '44,00,000')
      end

      expect(paragraph.text).to eq('Your title is Senior Account Executive at CTC 44,00,000.')
    end

    it 'leaves a lone unmatched "}}" untouched' do
      paragraph = paragraph_from('<w:r><w:t xml:space="preserve">Total is 100}} today</w:t></w:r>')

      expect(paragraph.text).to eq('Total is 100}} today')
    end

    it 'leaves a lone unmatched "{{" untouched' do
      paragraph = paragraph_from('<w:r><w:t xml:space="preserve">Value {{ is here</w:t></w:r>')

      expect(paragraph.text).to eq('Value {{ is here')
    end

    it 'leaves a balanced empty "{{}}" untouched' do
      paragraph = paragraph_from('<w:r><w:t xml:space="preserve">Ref {{}} end</w:t></w:r>')

      expect(paragraph.text).to eq('Ref {{}} end')
    end

    it 'consolidates a real token but leaves an author-typed stray "}}" in place' do
      paragraph = paragraph_from(
        '<w:r><w:t>{{offer.job_title}}</w:t></w:r>' \
        '<w:r><w:t xml:space="preserve"> and }} stray</w:t></w:r>'
      )
      paragraph.each_text_run { |run| run.substitute('{{offer.job_title}}', 'Engineer') }

      expect(paragraph.text).to eq('Engineer and }} stray')
    end

    # A "{{ spaced }}" pseudo-token is not a real merge field, but the gem's
    # consolidation regex still treats any {{...}} as a placeholder. It must be
    # kept atomic (never fragmented into an orphan brace) even when it shares a
    # boundary run with a real token; callers whose token regex excludes spaces
    # simply leave it in place.
    it 'keeps a spaced "{{ abcd.abc }}" intact when it shares a boundary run with a real token' do
      paragraph = paragraph_from(<<~XML)
        <w:r><w:t xml:space="preserve">Title {{</w:t></w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r><w:t>offer.job_title</w:t></w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r><w:t xml:space="preserve">}} note {{</w:t></w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r><w:t xml:space="preserve"> abcd.abc </w:t></w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r><w:t>}}.</w:t></w:r>
      XML

      expect(paragraph.text).to eq('Title {{offer.job_title}} note {{ abcd.abc }}.')

      paragraph.each_text_run { |run| run.substitute('{{offer.job_title}}', 'Engineer') }

      expect(paragraph.text).to eq('Title Engineer note {{ abcd.abc }}.')
    end

    # --- Normal / happy-path cases: prove the reconstruction preserves ordinary
    # behaviour (single-run tokens, plain splits, multiple tokens, no tokens). ---

    it 'leaves a paragraph with no placeholders unchanged' do
      expect(paragraph_from('<w:r><w:t xml:space="preserve">Just plain text here.</w:t></w:r>').text)
        .to eq('Just plain text here.')
    end

    it 'substitutes a placeholder wholly contained in a single run' do
      paragraph = paragraph_from('<w:r><w:t xml:space="preserve">Hello {{offer.name}}!</w:t></w:r>')

      expect(paragraph.text).to eq('Hello {{offer.name}}!')

      paragraph.each_text_run { |run| run.substitute('{{offer.name}}', 'Jane') }
      expect(paragraph.text).to eq('Hello Jane!')
    end

    it 'consolidates and substitutes a placeholder split across two runs' do
      paragraph = paragraph_from(
        '<w:r><w:t xml:space="preserve">Hello {{offer.</w:t></w:r>' \
        '<w:r><w:t xml:space="preserve">name}}!</w:t></w:r>'
      )

      expect(paragraph.text).to eq('Hello {{offer.name}}!')

      paragraph.each_text_run { |run| run.substitute('{{offer.name}}', 'Jane') }
      expect(paragraph.text).to eq('Hello Jane!')
    end

    it 'consolidates a placeholder split across three runs (start / middle / end)' do
      paragraph = paragraph_from(
        '<w:r><w:t xml:space="preserve">A {{</w:t></w:r>' \
        '<w:r><w:t>offer.name</w:t></w:r>' \
        '<w:r><w:t xml:space="preserve">}} B</w:t></w:r>'
      )

      expect(paragraph.text).to eq('A {{offer.name}} B')

      paragraph.each_text_run { |run| run.substitute('{{offer.name}}', 'X') }
      expect(paragraph.text).to eq('A X B')
    end

    it 'substitutes multiple placeholders that sit in separate runs' do
      paragraph = paragraph_from(
        '<w:r><w:t xml:space="preserve">Name </w:t></w:r>' \
        '<w:r><w:t>{{offer.name}}</w:t></w:r>' \
        '<w:r><w:t xml:space="preserve"> Role </w:t></w:r>' \
        '<w:r><w:t>{{offer.role}}</w:t></w:r>'
      )

      paragraph.each_text_run do |run|
        run.substitute('{{offer.name}}', 'Jane')
        run.substitute('{{offer.role}}', 'Engineer')
      end

      expect(paragraph.text).to eq('Name Jane Role Engineer')
    end

    it 'substitutes two adjacent placeholders with no text between them' do
      paragraph = paragraph_from(
        '<w:r><w:t>{{offer.a}}</w:t></w:r>' \
        '<w:r><w:t>{{offer.b}}</w:t></w:r>'
      )

      paragraph.each_text_run do |run|
        run.substitute('{{offer.a}}', '1')
        run.substitute('{{offer.b}}', '2')
      end

      expect(paragraph.text).to eq('12')
    end

    # Google Docs exports .docx WITHOUT Word's <w:proofErr> markers and splits
    # runs at its own style boundaries (explicit <w:rPr> per run). The algorithm
    # only reads <w:r> text, so it must handle a shared "}} ... {{" boundary run
    # the same way regardless of authoring tool.
    it 'fixes a shared boundary run in a Google Docs style paragraph (no proofErr)' do
      styled = lambda do |text|
        rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="Arial" w:hAnsi="Arial" w:cs="Arial"/>' \
          '<w:color w:val="000000"/><w:sz w:val="22"/></w:rPr>'
        "<w:r>#{rpr}<w:t xml:space=\"preserve\">#{text}</w:t></w:r>"
      end
      paragraph = paragraph_from(
        styled.call('Title {{') + styled.call('offer.job_title') +
        styled.call('}} at CTC {{') + styled.call('offer.salary_amount') + styled.call('}}.')
      )

      expect(paragraph.text).to eq('Title {{offer.job_title}} at CTC {{offer.salary_amount}}.')

      paragraph.each_text_run do |run|
        run.substitute('{{offer.job_title}}', 'Engineer')
        run.substitute('{{offer.salary_amount}}', '100k')
      end

      expect(paragraph.text).to eq('Title Engineer at CTC 100k.')
    end
  end
end
